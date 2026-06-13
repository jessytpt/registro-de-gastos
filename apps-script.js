// Pega este código en Google Apps Script (Extensiones → Apps Script)
// Luego: Implementar → Nueva implementación → Aplicación web
// Ejecutar como: Yo | Acceso: Cualquier persona

const SHEET_EXPENSES   = 'Gastos';
const SHEET_CATEGORIES = 'Categorías';

function doGet(e) {
  if (e.parameter.action === 'ping') {
    return json({ ok: true });
  }
  return json({ ok: false, error: 'Usa POST para sincronizar' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'sync') {
      syncExpenses(body.expenses || []);
      syncCategories(body.categories || []);
      return json({ ok: true });
    }

    return json({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

function syncExpenses(expenses) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_EXPENSES);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EXPENSES);
  }

  sheet.clearContents();

  const headers = ['ID', 'Fecha', 'Descripción', 'Categoría', 'Monto'];
  sheet.appendRow(headers);

  // Cabecera en negrita
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  expenses.forEach(exp => {
    sheet.appendRow([
      exp.id,
      exp.date,
      exp.description,
      exp.category,
      exp.amount,
    ]);
  });

  // Formato moneda en columna Monto (E)
  if (expenses.length > 0) {
    sheet.getRange(2, 5, expenses.length, 1).setNumberFormat('$#,##0.00');
  }

  // Ajustar ancho de columnas
  sheet.autoResizeColumns(1, headers.length);
}

function syncCategories(categories) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_CATEGORIES);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CATEGORIES);
  }

  sheet.clearContents();

  const headers = ['ID', 'Nombre', 'Ícono', 'Color', 'Presupuesto Mensual'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  categories.forEach(cat => {
    sheet.appendRow([
      cat.id,
      cat.label,
      cat.icon,
      cat.color,
      cat.budget || 0,
    ]);
  });

  if (categories.length > 0) {
    sheet.getRange(2, 5, categories.length, 1).setNumberFormat('$#,##0.00');
  }

  sheet.autoResizeColumns(1, headers.length);
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
