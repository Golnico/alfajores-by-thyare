function doPost(e) {
  // Log raw request for debugging
  console.log('Raw request:', JSON.stringify(e));

  let data = {};

  // Try to read data from form parameters (URL-encoded)
  try {
    if (e.parameter) {
      data = e.parameter;
      console.log('Form data received:', JSON.stringify(data));
    }
  } catch (err) {
    console.error('Error reading form data:', err);
  }

  // Parse items JSON if present
  let items = [];
  try {
    if (data.items) {
      items = JSON.parse(data.items);
    }
  } catch (err) {
    console.error('Error parsing items:', err);
  }

  // Get spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Add to Pedidos sheet
  const pedidosSheet = ss.getSheetByName('Pedidos');
  if (!pedidosSheet) {
    return ContentService.createTextOutput(JSON.stringify({
      'status': 'error',
      'message': 'Sheet "Pedidos" not found'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const timestamp = new Date();

  // Format items as string
  let itemsString = '';
  items.forEach(item => {
    itemsString += `${item.qty}x ${item.name}, `;
  });
  itemsString = itemsString.slice(0, -2); // Remove last comma

  // Add row to Pedidos
  pedidosSheet.appendRow([
    timestamp,                    // A: Fecha/Hora
    data.name || '',               // B: Cliente
    data.phone || '',               // C: Teléfono
    itemsString,                   // D: Items
    parseInt(data.total) || 0,     // E: Total
    'Pendiente',                   // F: Estado Pago
    '',                            // G: Método Pago
    data.notes || '',              // H: Notas
    'Pendiente',                   // I: Estado Entrega
    'Nuevo'                        // J: Estado Pedido
  ]);

  console.log('Row added to Pedidos');

  // 2. Update Stock sheet
  const stockSheet = ss.getSheetByName('Stock');
  if (stockSheet && items.length > 0) {
    const stockData = stockSheet.getDataRange().getValues();

    items.forEach(item => {
      for (let i = 1; i < stockData.length; i++) {
        if (stockData[i][0] === item.id) {
          stockData[i][2] = (stockData[i][2] || 0) + item.qty;
          stockData[i][3] = stockData[i][1] - stockData[i][2];
          break;
        }
      }
    });

    stockSheet.getRange(1, 1, stockData.length, stockData[0].length).setValues(stockData);
    console.log('Stock updated');
  }

  return ContentService.createTextOutput(JSON.stringify({
    'status': 'success',
    'message': 'Pedido registrado'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    'status': 'ok',
    'message': 'Alfajores by Thyare - API activa'
  })).setMimeType(ContentService.MimeType.JSON);
}
