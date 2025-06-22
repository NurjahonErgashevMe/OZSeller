const fieldsToShow = [
  'id',              // Product ID OZON
  'offer_id',        // Внутренний артикул
  'marketing_price', // Цена с учетом софинансирования Ozon
  'price'            // Исходная цена
];

const flatConfig = {
  formatDate: 'dateTime',
  onlyFields: fieldsToShow,
  excludeFields: [],
  noFormatFields: ['offer_id', 'id'],
};