function main() {
    // --- НАЛАШТУВАННЯ ---
    var SPREADSHEET_URL = 'ТУТ_ВАШЕ_ПОСИЛАННЯ_НА_GOOGLE_SHEET';
    var FEED_URLS = ['URL_UK'];             // Масив посилань на XML-фіди (наприклад: ['URL_UK', 'URL_RU']). Залиште порожнім [], якщо не потрібно
    var ANALYZE_DAYS = 30;                  // Скільки днів аналізувати (наприклад, 30)
    var IGNORE_LAST_DAYS = 3;               // Скільки останніх днів ігнорувати (щоб не враховувати затримку атрибуції конверсій)

    var COST_THRESHOLD = 600;               // Мінімальні витрати товару без конверсій, після яких він потрапляє у звіт
    var PRICE_COST_RATIO_PERCENT = 30;      // Відсоток витрат від ціни товару (наприклад, 30), після якого товар без конверсій зупиняється
    var TARGET_ROAS = 200;                  // Цільовий ROAS у відсотках (наприклад, 250), нижче якого товар вважається слабким
    var MIN_COST_FOR_ROAS_CHECK = 200;      // Мінімальні витрати для перевірки ROAS, щоб не оцінювати товари з малою статистикою
    var SUPER_ROAS_PERCENT = 250;           // ROAS, вище якого товар отримує мітку успішного (top_roas)
    var MIN_CONVERSIONS_FOR_SUPER_ROAS = 2; // Мінімум конверсій, щоб уникнути випадкового успіху від 1 продажу
    var MIN_CLICKS_THRESHOLD = 230;          // Мінімум кліків для правила "0 конверсій", щоб уникати хибних висновків
    var MIN_IMPRESSIONS_FOR_CTR = 1000;     // Мінімум показів для перевірки CTR, щоб не аналізувати товари з малим охопленням
    var LOW_CTR_THRESHOLD = 0.3;            // Мінімальний CTR у %, нижче цього значення товар позначається як проблемний

    // --- МІТКИ ДЛЯ SUPPLEMENTAL FEED ---
    var LABEL_TOP_ROAS = 'top_roas';
    var LABEL_NO_CONVERSIONS = 'no_index';
    var LABEL_LOW_ROAS = 'low_roas';
    var LABEL_LOW_CTR = 'low_ctr';

    // --- ЗАВАНТАЖЕННЯ КОДІВ МОВ ---
    var languageCodeMap = fetchLanguageCodes();

    // --- ПІДГОТОВКА ТАБЛИЦІ ---
    var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

    var sheetName = 'Аналіз';
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
        sheet.clear();
    } else {
        sheet = spreadsheet.insertSheet(sheetName);
    }

    var header = [
        'Product ID',
        'Мова',
        'Feed Label',
        'Витрати (грн)',
        'Ціна товару',
        '% витрат',
        'Покази',
        'Кліки',
        'CTR (%)',
        'Конверсії',
        'Цінність конв.',
        'Поточний ROAS',
        'Статус проблеми',
        'Мітка (custom_label_3)'
    ];
    sheet.appendRow(header);
    sheet.getRange(1, 1, 1, header.length)
        .setFontWeight('bold')
        .setBackground('#f0f0f0');

    // Кеш для аркушів фіду (щоб не створювати їх двічі)
    var feedSheetsMap = {};

    // --- ОТРИМАННЯ ЦІН З ФІДА ---
    var priceMap = fetchPricesFromFeeds(FEED_URLS);

    // --- ПІДГОТОВКА ДАТИ ---
    var timeZone = AdsApp.currentAccount().getTimeZone();
    var now = new Date();

    var endDateObj = new Date(now.getTime() - IGNORE_LAST_DAYS * 24 * 60 * 60 * 1000);
    var startDateObj = new Date(endDateObj.getTime() - (ANALYZE_DAYS - 1) * 24 * 60 * 60 * 1000);

    var endDateStr = Utilities.formatDate(endDateObj, timeZone, 'yyyy-MM-dd');
    var startDateStr = Utilities.formatDate(startDateObj, timeZone, 'yyyy-MM-dd');
    var dateRangeCondition = "BETWEEN '" + startDateStr + "' AND '" + endDateStr + "'";

    // --- ЗАПИТ ---
    var query =
        'SELECT segments.product_item_id, ' +
        'segments.product_language, ' +
        'segments.product_feed_label, ' +
        'metrics.cost_micros, metrics.impressions, metrics.clicks, ' +
        'metrics.conversions, metrics.conversions_value ' +
        'FROM shopping_performance_view ' +
        'WHERE segments.date ' + dateRangeCondition +
        ' AND metrics.impressions > 0';

    var rows = AdsApp.search(query);

    var totalRows = 0;
    var issueCount = 0;

    while (rows.hasNext()) {
        var row = rows.next();
        totalRows++;

        var cost = row.metrics.costMicros / 1000000;
        var impressions = row.metrics.impressions;
        var clicks = row.metrics.clicks;
        var conversions = row.metrics.conversions;
        var value = row.metrics.conversionsValue;
        var productId = row.segments.productItemId;
        var rawLanguage = row.segments.productLanguage || '';
        var language = languageCodeMap[rawLanguage] || rawLanguage.replace('languageConstants/', '');

        var feedLabel = row.segments.productFeedLabel || '';
        var lookupId = productId ? productId.toLowerCase() : '';

        var feedData = priceMap[lookupId];
        var price = feedData ? feedData.price : 0;
        var originalId = feedData ? feedData.originalId : productId;

        var currentRoas = cost > 0 ? (value / cost) : 0;
        var ctr = impressions > 0 ? (clicks / impressions) * 100 : 0;
        var costRatioPercent = price > 0 ? (cost / price) * 100 : 0;
        var issueStatus = '';
        var label = '';

        if (conversions >= MIN_CONVERSIONS_FOR_SUPER_ROAS && (currentRoas * 100) >= SUPER_ROAS_PERCENT) {
            issueStatus = '🟢 Супер-рентабельний: ROAS ' + (currentRoas * 100).toFixed(0) + '%';
            label = LABEL_TOP_ROAS;
        }
        else if (price > 0 && conversions == 0 && costRatioPercent > PRICE_COST_RATIO_PERCENT) {
            issueStatus = '🔴 Злив бюджету: витрати > ' + PRICE_COST_RATIO_PERCENT + '% від ціни (' + Math.round(costRatioPercent) + '%)';
            label = LABEL_NO_CONVERSIONS;
        }
        else if (conversions == 0 && cost > COST_THRESHOLD && clicks >= MIN_CLICKS_THRESHOLD) {
            issueStatus = '🔴 Злив бюджету: 0 конверсій при ' + clicks + ' кліках';
            label = LABEL_NO_CONVERSIONS;
        }
        else if (conversions > 0 && cost > MIN_COST_FOR_ROAS_CHECK && (currentRoas * 100) < TARGET_ROAS) {
            issueStatus = '🟡 Низька ефективність: ROAS ' + (currentRoas * 100).toFixed(0) + '% (ціль ' + TARGET_ROAS + '%)';
            label = LABEL_LOW_ROAS;
        }
        else if (impressions >= MIN_IMPRESSIONS_FOR_CTR && ctr < LOW_CTR_THRESHOLD) {
            issueStatus = '🔵 Низький CTR (' + ctr.toFixed(2) + '%) — перевір фід або ціну';
            label = LABEL_LOW_CTR;
        }

        if (issueStatus !== '') {
            issueCount++;

            sheet.appendRow([
                originalId,
                language,
                feedLabel,
                Math.round(cost),
                price > 0 ? Math.round(price) : '',
                price > 0 ? Math.round(costRatioPercent) + '%' : '',
                impressions,
                clicks,
                Math.round(ctr * 100) / 100,
                conversions,
                Math.round(value * 100) / 100,
                Math.round(currentRoas * 100) / 100,
                issueStatus,
                label
            ]);

            var lastRow = sheet.getLastRow();
            var range = sheet.getRange(lastRow, 1, 1, header.length);
            if (label === LABEL_TOP_ROAS) {
                range.setBackground('#D9EAD3');
            } else if (label === LABEL_NO_CONVERSIONS) {
                range.setBackground('#FFE0E0');
            } else if (label === LABEL_LOW_ROAS) {
                range.setBackground('#FFF3CD');
            } else if (label === LABEL_LOW_CTR) {
                range.setBackground('#E0F0FF');
            }

            // Динамічно отримуємо або створюємо аркуш для конкретної мови
            var safeLang = language ? language : 'unknown';
            var langFeedName = 'Feed_' + safeLang;

            if (!feedSheetsMap[langFeedName]) {
                var langSheet = spreadsheet.getSheetByName(langFeedName);
                if (!langSheet) {
                    langSheet = spreadsheet.insertSheet(langFeedName);
                } else {
                    langSheet.clear();
                }
                langSheet.appendRow(['id', 'custom_label_3']);
                langSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#d9ead3');
                feedSheetsMap[langFeedName] = langSheet;
            }

            feedSheetsMap[langFeedName].appendRow([originalId, label]);
        }
    }

    if (issueCount > 0) {
        var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, header.length);
        // Сортуємо спочатку за міткою (щоб top_roas був зверху), потім за витратами
        dataRange.sort([{ column: 14, ascending: false }, { column: 4, ascending: false }]);
    }

    Logger.log('✅ Аналіз завершено.');
    Logger.log('Оброблено товарів: ' + totalRows + ', проблемних: ' + issueCount);
    Logger.log('Лист аналізу: ' + sheetName);
    Logger.log('SupplementalFeed оновлено: ' + issueCount + ' товарів з мітками.');
}

function fetchPricesFromFeeds(feedUrls) {
    var priceMap = {};
    if (!feedUrls || feedUrls.length === 0) return priceMap;

    for (var u = 0; u < feedUrls.length; u++) {
        var feedUrl = feedUrls[u];
        if (!feedUrl) continue;

        Logger.log('🔄 Завантаження цін з фіда ' + (u + 1) + '...');
        try {
            var xml = UrlFetchApp.fetch(feedUrl).getContentText();

            // Використовуємо регулярні вирази замість XmlService, 
            // оскільки XmlService має суворі ліміти на розмір тексту (SAXException maxGeneralEntitySizeLimit)
            var itemRegex = /<item>([\s\S]*?)<\/item>/gi;
            var match;
            var count = 0;

            while ((match = itemRegex.exec(xml)) !== null) {
                var itemContent = match[1];

                // Шукаємо <g:id> та <g:price> всередині кожного <item>
                var idMatch = itemContent.match(/<g:id>([^<]+)<\/g:id>/i);
                var salePriceMatch = itemContent.match(/<g:sale_price>([^<]+)<\/g:sale_price>/i);
                var priceMatch = itemContent.match(/<g:price>([^<]+)<\/g:price>/i);

                if (idMatch && (salePriceMatch || priceMatch)) {
                    var id = idMatch[1].trim().toLowerCase();
                    var priceStr = salePriceMatch ? salePriceMatch[1] : priceMatch[1];

                    // Обробляємо коми як роздільник копійок
                    priceStr = priceStr.replace(',', '.');
                    priceStr = priceStr.replace(/[^0-9\.]/g, '');

                    var price = parseFloat(priceStr) || 0;
                    if (price > 0) {
                        priceMap[id] = {
                            price: price,
                            originalId: idMatch[1].trim()
                        };
                        count++;
                    }
                }
            }
            Logger.log('✅ Отримано ціни для ' + count + ' товарів з фіда ' + (u + 1) + '.');
        } catch (e) {
            Logger.log('❌ Помилка завантаження фіда ' + (u + 1) + ': ' + e.message);
        }
    }
    Logger.log('Усього унікальних товарів з цінами: ' + Object.keys(priceMap).length);
    return priceMap;
}

function fetchLanguageCodes() {
    var languageMap = {};
    try {
        var query = "SELECT language_constant.id, language_constant.code FROM language_constant";
        var rows = AdsApp.search(query);
        while (rows.hasNext()) {
            var row = rows.next();
            languageMap['languageConstants/' + row.languageConstant.id] = row.languageConstant.code;
        }
        Logger.log('✅ Отримано словник мов (' + Object.keys(languageMap).length + ' шт.)');
    } catch (e) {
        Logger.log('⚠️ Не вдалося завантажити коди мов. Використовуємо базовий словник. Помилка: ' + e.message);
        languageMap = {
            'languageConstants/1031': 'uk',
            'languageConstants/1014': 'ru',
            'languageConstants/1000': 'en'
        };
    }
    return languageMap;
}