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
  
  if (lastRow >= 6) {
    const numRows = lastRow - 5;
    const offerIdRange = sheet.getRange(6, 1, numRows, 1);
    const offerIdValues = offerIdRange.getValues();
    const artikulIdRange = sheet.getRange(6, 2, numRows, 1);
    const artikulIdValues = artikulIdRange.getValues();
    const finalPriceRange = sheet.getRange(6, 4, numRows, 1);
    const finalPriceValues = finalPriceRange.getValues();
    
    for (let i = 0; i < offerIdValues.length; i++) {
      const offerId = offerIdValues[i][0];
      if (offerId) {
        existingDataB[offerId] = artikulIdValues[i][0] || '';
        existingDataD[offerId] = finalPriceValues[i][0] || '';
      }
    }
  }

  const rows = productsFlatCollection.map(product => {
    const offerId = product.offer_id || '';
    const productId = product.id || '';
    const marketingPrice = product.marketing_price ? Number(product.marketing_price) : 0;
    const price = product.price ? Number(product.price) : 0;

    const finalPrice = Math.round(marketingPrice * (1 - walletPercent / 100));
    const cofinancingPercent = price ? parseFloat(((price - marketingPrice) / price * 100).toFixed(2)) : 0;

    const existingArtikulId = existingDataB[offerId] || '';
    const existingFinalPrice = existingDataD[offerId] || '';

    return [
      offerId,
      existingArtikulId,
      productId,
      existingFinalPrice,
      Math.round(marketingPrice) + ' ₽',
      cofinancingPercent + '%',
      Math.round(price) + ' ₽'
    ];
  });

  const productsInfoGrid = [header, ...rows];
  writeGridToTable(productsInfoGrid, sheet.getName());
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
    const finalPriceRange = sheet.getRange(6, 4, numRows, 1);
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

        const finalPrice = Math.round(marketingPrice * (1 - walletPercent / 100));
        const cofinancingPercent = price ? parseFloat(((price - marketingPrice) / price * 100).toFixed(2)) : 0;

        const existingArtikulId = existingDataB[offerId] || '';
        const existingFinalPrice = existingDataD[offerId] || '';

        rows.push([
          offerId,
          existingArtikulId,
          productId,
          existingFinalPrice,
          Math.round(marketingPrice) + ' ₽',
          cofinancingPercent + '%',
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
}

function onOpen(e) {
  addOzonInMenu();
}