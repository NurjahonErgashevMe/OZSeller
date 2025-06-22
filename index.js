function loadAndWriteAllProductsInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const apiKey = sheet.getRange("D1").getValue();
  const clientId = sheet.getRange("D2").getValue();
  const walletPercent = parseFloat(sheet.getRange("D3").getValue()) || 0;

  const keys = {
    clientId: Number(clientId),
    apiKey: String(apiKey)
  };
  const client = OzonApi.client(keys);

  const productsIdsListResponse = client.productApiGetProductList({
    filter: {
      visibility: 'ALL',
    },
    limit: 1000,
  });

  const productsOfferIds = productsIdsListResponse.result.items.map((item) => item.offer_id);

  const productsInfoListResponse = client.productApiGetProductInfoListV3({
    offer_id: productsOfferIds
  });

  const productsFlatCollection = productsInfoListResponse.items
    .map(product => flatter(product, flatConfig));

  if (!productsFlatCollection || productsFlatCollection.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Не получено данных о товарах из API', 'Ошибка');
    return;
  }

  const header = [
    'Внутренний артикул',
    'Артикул ID OZON',
    'Product ID OZON',
    'Конечная цена (с OZON кошельком)',
    'Цена с учетом софинансирования OZON',
    'Процент софинансирования',
    'Цена исходная (из личного кабинета)'
  ];

  const lastRow = sheet.getLastRow();
  const existingDataB = {};
  const existingDataD = {};
  const existingFormulasB = {}; // Сохраняем формулы для столбца B
  const existingFormulasD = {}; // Сохраняем формулы для столбца D
  const existingOrder = []; // Сохраняем порядок строк
  let existingFormatting = {}; // Сохраняем форматирование
  let columnFormatting = {}; // Форматирование столбцов
  
  if (lastRow >= 6) {
    const numRows = lastRow - 5;
    const dataRange = sheet.getRange(6, 1, numRows, 7); // Весь диапазон данных
    
    // Сохраняем форматирование
    const backgrounds = dataRange.getBackgrounds();
    const fontColors = dataRange.getFontColors();
    const fontFamilies = dataRange.getFontFamilies();
    const fontSizes = dataRange.getFontSizes();
    const fontWeights = dataRange.getFontWeights();
    const horizontalAlignments = dataRange.getHorizontalAlignments();
    const verticalAlignments = dataRange.getVerticalAlignments();
    const numberFormats = dataRange.getNumberFormats();
    
    existingFormatting = {
      backgrounds: backgrounds,
      fontColors: fontColors,
      fontFamilies: fontFamilies,
      fontSizes: fontSizes,
      fontWeights: fontWeights,
      horizontalAlignments: horizontalAlignments,
      verticalAlignments: verticalAlignments,
      numberFormats: numberFormats
    };
    
    // Определяем форматирование для каждого столбца (берем из первой строки как шаблон)
    if (numRows > 0) {
      for (let col = 0; col < 7; col++) {
        columnFormatting[col] = {
          horizontalAlignment: horizontalAlignments[0][col],
          verticalAlignment: verticalAlignments[0][col],
          fontFamily: fontFamilies[0][col],
          fontSize: fontSizes[0][col],
          numberFormat: numberFormats[0][col]
        };
      }
    }
    
    const offerIdRange = sheet.getRange(6, 1, numRows, 1);
    const offerIdValues = offerIdRange.getValues();
    const artikulIdRange = sheet.getRange(6, 2, numRows, 1);
    const artikulIdValues = artikulIdRange.getValues();
    const artikulIdFormulas = artikulIdRange.getFormulas(); // Получаем формулы для B
    const finalPriceRange = sheet.getRange(6, 4, numRows, 1);
    const finalPriceValues = finalPriceRange.getValues();
    const finalPriceFormulas = finalPriceRange.getFormulas(); // Получаем формулы для D
    
    for (let i = 0; i < offerIdValues.length; i++) {
      const offerId = offerIdValues[i][0];
      if (offerId) {
        existingDataB[offerId] = artikulIdValues[i][0] || '';
        existingDataD[offerId] = finalPriceValues[i][0] || '';
        // Сохраняем формулы если они есть
        if (artikulIdFormulas[i][0]) {
          existingFormulasB[offerId] = artikulIdFormulas[i][0];
        }
        if (finalPriceFormulas[i][0]) {
          existingFormulasD[offerId] = finalPriceFormulas[i][0];
        }
        existingOrder.push(offerId);
      } else {
        existingOrder.push(null); // Пустая строка
      }
    }
  }

  // Создаем карту продуктов для быстрого поиска
  const productsMap = new Map();
  productsFlatCollection.forEach(product => {
    if (product.offer_id) {
      productsMap.set(product.offer_id, product);
    }
  });

  const rows = [];
  
  // Если есть существующий порядок, используем его
  if (existingOrder.length > 0) {
    existingOrder.forEach(offerId => {
      if (offerId === null) {
        // Пустая строка
        rows.push(['', '', '', '', '', '', '']);
      } else {
        const product = productsMap.get(offerId);
        if (product) {
          const productId = product.id || '';
          const marketingPrice = product.marketing_price ? Number(product.marketing_price) : 0;
          const price = product.price ? Number(product.price) : 0;
          
          const cofinancingPercent = price ? parseFloat(((price - marketingPrice) / price * 100).toFixed(2)) : 0;
          
          const existingArtikulId = existingDataB[offerId] || '';
          const existingFinalPrice = existingDataD[offerId] || '';
          
          rows.push([
            offerId,
            existingArtikulId,
            productId,
            existingFinalPrice,
            Math.round(marketingPrice) + ' ₽',
            cofinancingPercent.toString().replace('.', ',') + '%',
            Math.round(price) + ' ₽'
          ]);
        } else {
          // Товар не найден в новых данных, но оставляем строку
          const existingArtikulId = existingDataB[offerId] || '';
          const existingFinalPrice = existingDataD[offerId] || '';
          
          rows.push([
            offerId,
            existingArtikulId,
            '',
            existingFinalPrice,
            '',
            '',
            ''
          ]);
        }
      }
    });
    
    // Добавляем новые товары, которых не было в существующих данных
    productsFlatCollection.forEach(product => {
      const offerId = product.offer_id || '';
      if (offerId && !existingDataB.hasOwnProperty(offerId)) {
        const productId = product.id || '';
        const marketingPrice = product.marketing_price ? Number(product.marketing_price) : 0;
        const price = product.price ? Number(product.price) : 0;
        
        const cofinancingPercent = price ? parseFloat(((price - marketingPrice) / price * 100).toFixed(2)) : 0;
        
        rows.push([
          offerId,
          '',
          productId,
          '',
          Math.round(marketingPrice) + ' ₽',
          cofinancingPercent.toString().replace('.', ',') + '%',
          Math.round(price) + ' ₽'
        ]);
      }
    });
  } else {
    // Если нет существующих данных, создаем новые строки
    productsFlatCollection.forEach(product => {
      const offerId = product.offer_id || '';
      const productId = product.id || '';
      const marketingPrice = product.marketing_price ? Number(product.marketing_price) : 0;
      const price = product.price ? Number(product.price) : 0;
      
      const cofinancingPercent = price ? parseFloat(((price - marketingPrice) / price * 100).toFixed(2)) : 0;
      
      rows.push([
        offerId,
        '',
        productId,
        '',
        Math.round(marketingPrice) + ' ₽',
        cofinancingPercent.toString().replace('.', ',') + '%',
        Math.round(price) + ' ₽'
      ]);
    });
  }

  // Записываем данные, но сохраняем формулы в столбцах B и D
  const productsInfoGrid = [header, ...rows];
  
  // Записываем заголовок
  const headerRange = sheet.getRange(5, 1, 1, 7);
  headerRange.setValues([header]);
  
  // Записываем данные построчно, сохраняя формулы в B и D
  if (rows.length > 0) {
    for (let i = 0; i < rows.length; i++) {
      const rowIndex = i + 6; // Строка в таблице (начинаем с 6-й)
      const rowData = rows[i];
      const offerId = rowData[0];
      
      // Для всех столбцов
      for (let col = 0; col < 7; col++) {
        const cellRange = sheet.getRange(rowIndex, col + 1);
        const isFormulaColumn = col === 1 || col === 3; // B и D
        
        if (isFormulaColumn) {
          // Восстанавливаем сохраненные формулы
          if (col === 1 && existingFormulasB[offerId]) {
            cellRange.setFormula(existingFormulasB[offerId]);
          } 
          else if (col === 3 && existingFormulasD[offerId]) {
            cellRange.setFormula(existingFormulasD[offerId]);
          }
          else {
            // Если формулы нет, записываем значение
            cellRange.setValue(rowData[col]);
          }
        } else {
          // Для остальных столбцов всегда записываем значение
          cellRange.setValue(rowData[col]);
        }
      }
    }
  }
  
  // Восстанавливаем форматирование
  if (rows.length > 0) {
    const currentLastRow = sheet.getLastRow();
    if (currentLastRow >= 6) {
      const currentNumRows = currentLastRow - 5;
      const restoreRange = sheet.getRange(6, 1, currentNumRows, 7);
      
      // Если есть сохраненное форматирование, используем его
      if (existingFormatting.backgrounds) {
        const numExistingRows = Math.min(currentNumRows, existingFormatting.backgrounds.length);
        
        if (numExistingRows > 0) {
          const existingRange = sheet.getRange(6, 1, numExistingRows, 7);
          // Обрезаем массивы форматирования под количество существующих строк
          const trimmedBackgrounds = existingFormatting.backgrounds.slice(0, numExistingRows);
          const trimmedFontColors = existingFormatting.fontColors.slice(0, numExistingRows);
          const trimmedFontFamilies = existingFormatting.fontFamilies.slice(0, numExistingRows);
          const trimmedFontSizes = existingFormatting.fontSizes.slice(0, numExistingRows);
          const trimmedFontWeights = existingFormatting.fontWeights.slice(0, numExistingRows);
          const trimmedHorizontalAlignments = existingFormatting.horizontalAlignments.slice(0, numExistingRows);
          const trimmedVerticalAlignments = existingFormatting.verticalAlignments.slice(0, numExistingRows);
          const trimmedNumberFormats = existingFormatting.numberFormats.slice(0, numExistingRows);
          
          // Восстанавливаем форматирование для существующих строк
          existingRange.setBackgrounds(trimmedBackgrounds);
          existingRange.setFontColors(trimmedFontColors);
          existingRange.setFontFamilies(trimmedFontFamilies);
          existingRange.setFontSizes(trimmedFontSizes);
          existingRange.setFontWeights(trimmedFontWeights);
          existingRange.setHorizontalAlignments(trimmedHorizontalAlignments);
          existingRange.setVerticalAlignments(trimmedVerticalAlignments);
          existingRange.setNumberFormats(trimmedNumberFormats);
        }
        
        // Применяем форматирование столбцов к новым строкам
        if (currentNumRows > numExistingRows && Object.keys(columnFormatting).length > 0) {
          const newRowsCount = currentNumRows - numExistingRows;
          const newRowsRange = sheet.getRange(6 + numExistingRows, 1, newRowsCount, 7);
          
          // Применяем форматирование по столбцам к новым строкам
          for (let col = 1; col <= 7; col++) {
            const colIndex = col - 1;
            if (columnFormatting[colIndex]) {
              const columnRange = sheet.getRange(6 + numExistingRows, col, newRowsCount, 1);
              
              if (columnFormatting[colIndex].horizontalAlignment) {
                columnRange.setHorizontalAlignment(columnFormatting[colIndex].horizontalAlignment);
              }
              if (columnFormatting[colIndex].verticalAlignment) {
                columnRange.setVerticalAlignment(columnFormatting[colIndex].verticalAlignment);
              }
              if (columnFormatting[colIndex].fontFamily) {
                columnRange.setFontFamily(columnFormatting[colIndex].fontFamily);
              }
              if (columnFormatting[colIndex].fontSize) {
                columnRange.setFontSize(columnFormatting[colIndex].fontSize);
              }
              if (columnFormatting[colIndex].numberFormat) {
                columnRange.setNumberFormat(columnFormatting[colIndex].numberFormat);
              }
            }
          }
        }
      }
    }
  }
}

function loadAndWriteProductsByIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const apiKey = sheet.getRange("D1").getValue();
  const clientId = sheet.getRange("D2").getValue();
  const walletPercent = parseFloat(sheet.getRange("D3").getValue()) || 0;

  const keys = {
    clientId: Number(clientId),
    apiKey: String(apiKey)
  };
  const client = OzonApi.client(keys);

  const lastRow = sheet.getLastRow();
  const productIds = [];
  
  if (lastRow < 6) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Нет данных в таблице для обработки. Добавьте Product ID в столбец C начиная с 6-й строки.', 'Ошибка');
    return;
  }
  
  const numRows = lastRow - 5;
  const productIdRange = sheet.getRange(6, 3, numRows, 1);
  const productIdValues = productIdRange.getValues();
  
  // Сохраняем форматирование перед парсингом
  const dataRange = sheet.getRange(6, 1, numRows, 7); // Весь диапазон данных
  const existingFormatting = {
    backgrounds: dataRange.getBackgrounds(),
    fontColors: dataRange.getFontColors(),
    fontFamilies: dataRange.getFontFamilies(),
    fontSizes: dataRange.getFontSizes(),
    fontWeights: dataRange.getFontWeights(),
    horizontalAlignments: dataRange.getHorizontalAlignments(),
    verticalAlignments: dataRange.getVerticalAlignments(),
    numberFormats: dataRange.getNumberFormats()
  };
  
  // Определяем форматирование для каждого столбца
  let columnFormatting = {};
  if (numRows > 0) {
    for (let col = 0; col < 7; col++) {
      columnFormatting[col] = {
        horizontalAlignment: existingFormatting.horizontalAlignments[0][col],
        verticalAlignment: existingFormatting.verticalAlignments[0][col],
        fontFamily: existingFormatting.fontFamilies[0][col],
        fontSize: existingFormatting.fontSizes[0][col],
        numberFormat: existingFormatting.numberFormats[0][col]
      };
    }
  }
  
  // Сохраняем формулы из столбцов B и D перед парсингом
  const artikulIdRange = sheet.getRange(6, 2, numRows, 1);
  const artikulIdFormulas = artikulIdRange.getFormulas();
  const finalPriceRange = sheet.getRange(6, 4, numRows, 1);
  const finalPriceFormulas = finalPriceRange.getFormulas();
  
  const productIdsWithPositions = [];
  productIdValues.forEach((row, index) => {
    const productId = row[0];
    if (productId && productId !== '' && !isNaN(productId)) {
      productIdsWithPositions.push({
        id: Number(productId),
        originalPosition: index,
        rowNumber: index + 6
      });
      productIds.push(Number(productId));
    }
  });
  
  if (productIds.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Не найдено валидных Product ID в столбце C', 'Ошибка');
    return;
  }

  const productsInfoListResponse = client.productApiGetProductInfoListV3({
    product_id: productIds.map((i) => i.toString())
  });

  const productsFlatCollection = productsInfoListResponse.items
    .map(product => flatter(product, flatConfig));

  const productsMap = new Map();
  productsFlatCollection.forEach(product => {
    if (product.id) {
      productsMap.set(Number(product.id), product);
    }
  });

  const header = [
    'Внутренний артикул',
    'Артикул ID OZON',
    'Product ID OZON',
    'Конечная цена (с OZON кошельком)',
    'Цена с учетом софинансирования OZON',
    'Процент софинансирования',
    'Цена исходная (из личного кабинета)'
  ];

  const existingDataB = {};
  const existingDataD = {};
  
  if (lastRow >= 6) {
    const offerIdRange = sheet.getRange(6, 1, numRows, 1);
    const offerIdValues = offerIdRange.getValues();
    const artikulIdValues = artikulIdRange.getValues();
    const finalPriceValues = finalPriceRange.getValues();
    
    for (let i = 0; i < offerIdValues.length; i++) {
      const offerId = offerIdValues[i][0];
      if (offerId) {
        existingDataB[offerId] = artikulIdValues[i][0] || '';
        existingDataD[offerId] = finalPriceValues[i][0] || '';
      }
    }
  }

  const rows = [];
  
  for (let i = 0; i < numRows; i++) {
    const cellValue = productIdValues[i][0];
    
    if (cellValue && cellValue !== '' && !isNaN(cellValue)) {
      const productId = Number(cellValue);
      const product = productsMap.get(productId);
      
      if (product) {
        const offerId = product.offer_id || '';
        const marketingPrice = product.marketing_price ? Number(product.marketing_price) : 0;
        const price = product.price ? Number(product.price) : 0;

        const cofinancingPercent = price ? parseFloat(((price - marketingPrice) / price * 100).toFixed(2)) : 0;

        const existingArtikulId = existingDataB[offerId] || '';
        const existingFinalPrice = existingDataD[offerId] || '';

        rows.push([
          offerId,
          existingArtikulId,
          productId,
          existingFinalPrice,
          Math.round(marketingPrice) + ' ₽',
          cofinancingPercent.toString().replace('.', ',') + '%',
          Math.round(price) + ' ₽'
        ]);
      } else {
        const existingArtikulId = '';
        const existingFinalPrice = '';
        
        rows.push([
          '',
          existingArtikulId,
          productId,
          existingFinalPrice,
          '',
          '',
          ''
        ]);
      }
    } else {
      rows.push(['', '', cellValue || '', '', '', '', '']);
    }
  }

  // Записываем данные, сохраняя формулы в B и D
  const productsInfoGrid = [header, ...rows];
  
  // Записываем заголовок
  const headerRange = sheet.getRange(5, 1, 1, 7);
  headerRange.setValues([header]);
  
  // Записываем данные построчно
  if (rows.length > 0) {
    for (let i = 0; i < rows.length; i++) {
      const rowIndex = i + 6; // Строка в таблице (начинаем с 6-й)
      const rowData = rows[i];
      
      for (let col = 0; col < 7; col++) {
        const cellRange = sheet.getRange(rowIndex, col + 1);
        const isFormulaColumn = col === 1 || col === 3; // B и D
        
        if (isFormulaColumn) {
          // Проверяем, была ли формула в исходных данных
          if (artikulIdFormulas[i][0] && col === 1) {
            cellRange.setFormula(artikulIdFormulas[i][0]);
          } 
          else if (finalPriceFormulas[i][0] && col === 3) {
            cellRange.setFormula(finalPriceFormulas[i][0]);
          }
          else {
            // Если формулы не было, записываем значение
            cellRange.setValue(rowData[col]);
          }
        } else {
          // Для остальных столбцов всегда записываем значение
          cellRange.setValue(rowData[col]);
        }
      }
    }
  }
  
  // Восстанавливаем форматирование с учетом форматирования столбцов
  if (existingFormatting.backgrounds && rows.length > 0) {
    const restoreRange = sheet.getRange(6, 1, rows.length, 7);
    
    // Обрезаем массивы форматирования под актуальное количество строк
    const trimmedBackgrounds = existingFormatting.backgrounds.slice(0, rows.length);
    const trimmedFontColors = existingFormatting.fontColors.slice(0, rows.length);
    const trimmedFontFamilies = existingFormatting.fontFamilies.slice(0, rows.length);
    const trimmedFontSizes = existingFormatting.fontSizes.slice(0, rows.length);
    const trimmedFontWeights = existingFormatting.fontWeights.slice(0, rows.length);
    const trimmedHorizontalAlignments = existingFormatting.horizontalAlignments.slice(0, rows.length);
    const trimmedVerticalAlignments = existingFormatting.verticalAlignments.slice(0, rows.length);
    const trimmedNumberFormats = existingFormatting.numberFormats.slice(0, rows.length);
    
    // Восстанавливаем форматирование
    restoreRange.setBackgrounds(trimmedBackgrounds);
    restoreRange.setFontColors(trimmedFontColors);
    restoreRange.setFontFamilies(trimmedFontFamilies);
    restoreRange.setFontSizes(trimmedFontSizes);
    restoreRange.setFontWeights(trimmedFontWeights);
    restoreRange.setHorizontalAlignments(trimmedHorizontalAlignments);
    restoreRange.setVerticalAlignments(trimmedVerticalAlignments);
    restoreRange.setNumberFormats(trimmedNumberFormats);
    
    // Дополнительно применяем форматирование столбцов ко всем строкам
    if (Object.keys(columnFormatting).length > 0) {
      for (let col = 1; col <= 7; col++) {
        const colIndex = col - 1;
        if (columnFormatting[colIndex]) {
          const columnRange = sheet.getRange(6, col, rows.length, 1);
          
          if (columnFormatting[colIndex].horizontalAlignment) {
            columnRange.setHorizontalAlignment(columnFormatting[colIndex].horizontalAlignment);
          }
          if (columnFormatting[colIndex].verticalAlignment) {
            columnRange.setVerticalAlignment(columnFormatting[colIndex].verticalAlignment);
          }
        }
      }
    }
  }
}
