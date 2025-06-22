/**
 * @description Преобразует номер столбца в буквенное обозначение
 * @param {number} col - номер столбца
 * @returns {string}  Буквенное обозначение колонки, соответствующее номеру.
 */
function getColSymbol(col) {
    const loopCount = Math.floor(col / 26);
    const colPosition = col % 26;
    let colSymbol;

    if (colPosition && loopCount) {
        colSymbol = String.fromCodePoint(64 + loopCount) + String.fromCodePoint(64 + colPosition);
    } else if (!colPosition && loopCount) {
        colSymbol = (String.fromCodePoint(64 + loopCount - 1) + String.fromCodePoint(65 + 25)).replace('@', '');
    } else {
        colSymbol = String.fromCodePoint(64 + colPosition);
    }

    return colSymbol;
}

/**
 * @description build A1 style range by coords (ex. "'orders'!A2:AX4")
 * @param {number} startCol
 * @param {number} startRow
 * @param {number} endCol
 * @param {number} endRow
 * @param {string} [sheetName]
 * @returns {string} range in  A1 style
 */
function buildA1String(startCol, startRow, endCol, endRow, sheetName) {
    const range = `${sheetName ? sheetName + '!' : ''}${getColSymbol(startCol)}${startRow}:${getColSymbol(
        endCol,
    )}${endRow}`;
    return range;
}

/**
 * Преобразует массив объектов в массив массивов (строк). Если не указаны свои заголовки,
 * то в качестве заголовков используются ключи первого в массиве объекта.
 * Значения в строках расставляются в последовательности соответствующей заголовку.
 * Если флаг options.includeHeader установлен в значение true, то первой строкой является значение заголовков
 * @param {Array<object>} collection - массив объектов, которые будут преобразованы в строки
 * @param {{ ownHeaders?: Array<string>; includeHeader: boolean }} options
 * @param {string[]} [options.ownHeaders] - массив, последовательность заголовков
 * @param {boolean} [options.includeHeader] - включить строку заголовков в результат
 * @returns {Array<Array<(string|number|boolean)>>}  массив массивов (строк)
 */
function collectionToGrid(collection, options) {
    const includeHeader = options?.includeHeader === true ? true : false;
    const headers = options?.ownHeaders || Object.keys(collection[0]);
    const grid = collection.map((obj) => {
        const row = [];
        for (const header of headers) {
            row.push(obj[header] || '');
        }
        return row;
    });

    if (includeHeader === true) {
        return [headers, ...grid];
    }

    return grid;
}

/**
 * Функция для "сплющивания" объектов. Выводит все вложенные поля на верхний уровень
 * @date 29.03.2023 - 11:04:27
 * @param {Object} data - object for flat
 * @param {{
    data: object;
    prefix?: string;
    customKey?: string;
    formatDate?: 'date' | 'dateTime';
    onlyFields?: string[];
    excludeFields?: string[];
    noFormatFields?: string[];
}} [config]
 * @returns {{[key: string]: number | string | Date | undefined;}}
 */
function flatter(data,config){

    function getType(obj) {
        if (obj === undefined) {
            throw new Error('no obj was passed for detect his type.');
        }
        const objType = Object.prototype.toString.call(obj);
        return objType.match(/\s(?<type>\w+)\]/)?.groups?.type;
    }

    function isDate(string) {
        return !isNaN(Date.parse(string));
    }

    if(!config){
       config = {};
    }

    const prefix = config.prefix || '';
    const { customKey, formatDate, onlyFields, excludeFields, noFormatFields } = config;
    let flatObj = {};
    for (const key in data) {
        if (onlyFields?.length && !onlyFields?.includes(key)) {
            continue;
        }

        if (excludeFields?.length && excludeFields?.includes(key)) {
            continue;
        }

        const valueType = getType(data[key]);

        if (noFormatFields?.includes(key) || valueType === 'Boolean') {
            flatObj[`${prefix}${customKey || key}`] = data[key];
            continue;
        }

        if (data[key] === undefined) {
            flatObj[`${prefix}${customKey || key}`] = undefined;
            continue;
        }

        if (formatDate && (valueType == 'String' || valueType == 'Date') && isDate(data[key])) {
            flatObj[`${prefix}${customKey || key}`] =
                formatDate == 'date'
                    ? new Date(data[key]).toLocaleDateString('ru')
                    : new Date(data[key]).toLocaleString('ru');
            continue;
        }

        if (Array.isArray(data[key])) {
            data[key].forEach((v, i) => {
                if (getType(v) == 'Object') {
                    Object.assign(flatObj, flatter(v, { prefix: `${key}[${i}].` }));
                    return;
                }
                if (Array.isArray(v)) {
                    const flatArr = {};
                    v.forEach((arrV, ii) => {
                        flatArr[`${prefix || key}[${i}][${ii}]`] = arrV;
                    });
                    Object.assign(flatObj, flatArr);
                    return;
                }
                return (flatObj[`${prefix}${customKey || key}[${i}]`] = v);
            });
            continue;
        }

        if (valueType == 'Object') {
            if (!Object.keys(data[key]).length) {
                flatObj[`${prefix}${customKey || key}`] = undefined;
                continue;
            }
            Object.assign(
                flatObj,
                flatter(data[key], {
                    prefix: `${prefix || key}.`,
                    formatDate,
                }),
            );
        }

        if (valueType == 'String' && data[key].length && !isNaN(data[key])) {
            flatObj[`${prefix}${customKey || key}`] = Number(data[key]);
            continue;
        }

        if (['String', 'Number'].includes(valueType)) {
            flatObj[`${prefix}${customKey || key}`] = data[key];
            continue;
        }
    }

    return flatObj;
}

/**
 * Очищает и затем записывает данные в указанный лист таблицы начиная с 6-й строки
 * @param {Array<Array<string|number|undefined>>} grid - массив массивов(строк) с данными для записи
 * @param {string} sheetName - Имя листа, в который данные будут записываться
 */
function writeGridToTable(grid, sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
        throw new Error(`Лист с именем "${sheetName}" не найден`);
    }
    
    // Проверяем, есть ли данные для записи
    if (!grid || grid.length === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Нет данных для записи', 'Предупреждение');
        return false;
    }
    
    // Определяем количество колонок и строк в данных
    const numCols = grid[0].length;
    const numRows = grid.length;
    
    // Очищаем старые данные начиная с 6-й строки
    const lastRow = sheet.getLastRow();
    if (lastRow >= 6) {
        const clearRange = sheet.getRange(6, 1, lastRow - 5, numCols);
        clearRange.clear();
    }
    
    // Записываем новые данные начиная с 6-й строки (без заголовков, так как они уже есть в строке 5)
    const dataToWrite = grid.slice(1); // Убираем заголовки из данных
    
    // Проверяем, есть ли данные после удаления заголовков
    if (dataToWrite.length === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Нет данных для записи (только заголовки)', 'Предупреждение');
        return false;
    }
    
    // Записываем данные
    const writeRange = sheet.getRange(6, 1, dataToWrite.length, numCols);
    writeRange.setValues(dataToWrite);
    
    // Показываем уведомление
    SpreadsheetApp.getActiveSpreadsheet().toast('Данные успешно записаны в ' + sheetName, 'Ozon Products');
    return true;
}

/** Показывает сообщение */
function showMsg(){
    const hello_ = `
    <div
    style="background-color: lightgreen; border-radius: 10px; padding: 10px; text-align: center; font-family: cursive;">
    Приятного использования!<br>
    Поддержка и ответы на вопросы <a href="https://t.me/google_sheets_pro">здесь</a>
    </div>`;

    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutput(hello_)
    .setHeight(80);
    ui.showModalDialog(html, 'GoogleSheets.ru');
}

/** Добавляет пункты в меню */
function addOzonInMenu() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Ozon')
      .addItem('Загрузить все товары из ЛК', 'loadAndWriteAllProductsInfo')
      .addItem('Загрузить товары по Product ID', 'loadAndWriteProductsByIds')
      .addToUi();
}