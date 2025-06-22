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
  const existingFormulas = {}; // Сохраняем формулы
  const existingOrder = []; // Сохраняем порядок строк
  
  if (lastRow >= 6) {
    const numRows = lastRow - 5;
    const offerIdRange = sheet.getRange(6, 1, numRows, 1);
    const offerIdValues = offerIdRange.getValues();
    const artikulIdRange = sheet.getRange(6, 2, numRows, 1);
    const artikulIdValues = artikulIdRange.getValues();
    const finalPriceRange = sheet.getRange(6, 4, numRows, 1);
    const finalPriceValues = finalPriceRange.getValues();
    const finalPriceFormulas = finalPriceRange.getFormulas(); // Получаем формулы
    
    for (let i = 0; i < offerIdValues.length; i++) {
      const offerId = offerIdValues[i][0];
      if (offerId) {
        existingDataB[offerId] = artikulIdValues[i][0] || '';
        existingDataD[offerId] = finalPriceValues[i][0] || '';
        // Сохраняем формулу если она есть
        if (finalPriceFormulas[i][0]) {
          existingFormulas[offerId] = finalPriceFormulas[i][0];
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
            cofinancingPercent.toString().replace('.', ',') + '%', // Число с % и запятой вместо точки
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
          cofinancingPercent.toString().replace('.', ',') + '%', // Число с % и запятой вместо точки
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
      
      const cofinancingPercent = price ? Math.round((price - marketingPrice) / price * 100) : 0; // Округляем до целого
      
      rows.push([
        offerId,
        '',
        productId,
        '',
        Math.round(marketingPrice) + ' ₽',
        cofinancingPercent, // Теперь это число, а не строка с %
        Math.round(price) + ' ₽'
      ]);
    });
  }

  const productsInfoGrid = [header, ...rows];
  writeGridToTable(productsInfoGrid, sheet.getName());
  
  // Восстанавливаем формулы после записи данных
  if (Object.keys(existingFormulas).length > 0) {
    const currentLastRow = sheet.getLastRow();
    if (currentLastRow >= 6) {
      const currentNumRows = currentLastRow - 5;
      const currentOfferIdRange = sheet.getRange(6, 1, currentNumRows, 1);
      const currentOfferIdValues = currentOfferIdRange.getValues();
      
      for (let i = 0; i < currentOfferIdValues.length; i++) {
        const offerId = currentOfferIdValues[i][0];
        if (offerId && existingFormulas[offerId]) {
          const cellRange = sheet.getRange(6 + i, 4); // Столбец D (4)
          cellRange.setFormula(existingFormulas[offerId]);
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
  
  // Сохраняем формулы из столбца D перед парсингом
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
    const artikulIdRange = sheet.getRange(6, 2, numRows, 1);
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
          cofinancingPercent + '%', // Теперь это число с % и двумя знаками после запятой
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

  const productsInfoGrid = [header, ...rows];
  writeGridToTable(productsInfoGrid, sheet.getName());
  
  // Восстанавливаем формулы после записи данных
  for (let i = 0; i < finalPriceFormulas.length; i++) {
    if (finalPriceFormulas[i][0]) {
      const cellRange = sheet.getRange(6 + i, 4); // Столбец D (4)
      cellRange.setFormula(finalPriceFormulas[i][0]);
    }
  }
}

function onOpen(e) {
  addOzonInMenu();
}