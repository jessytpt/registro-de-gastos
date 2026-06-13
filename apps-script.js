// Pega este código en Google Apps Script (Extensiones → Apps Script)
// Luego: Implementar → Nueva implementación → Aplicación web
// Ejecutar como: Yo | Acceso: Cualquier persona
// IMPORTANTE: cada vez que modifiques este código debes crear una NUEVA implementación

// *** CAMBIA ESTO POR TU CONTRASEÑA ***
const SECRET_KEY = 'mi-contraseña-secreta';

const SHEET_EXPENSES   = 'Gastos';
const SHEET_CATEGORIES = 'Categorías';

function doGet(e) {
  const action = e.parameter.action;
  const key    = e.parameter.key;

  if (!checkKey(key)) return json({ ok: false, error: 'No autorizado' });

  if (action === 'ping') {
    return json({ ok: true });
  }

  if (action === 'load') {
    const expenses   = readExpenses();
    const categories = readCategories();
    return json({ ok: true, expenses, categories });
  }

  return json({ ok: false, error: 'Acción no reconocida' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (!checkKey(body.key)) return json({ ok: false, error: 'No autorizado' });

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

function checkKey(key) {
  return key === SECRET_KEY;
}

// ── Leer ──────────────────────────────────────────────

function readExpenses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EXPENSES);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows
    .filter(r => r[0] !== '')
    .map(r => ({
      id:          Number(r[0]),
      date:        r[1],
      description: r[2],
      category:    r[3],
      amount:      Number(r[4]),
    }));
}

function readCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATEGORIES);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows
    .filter(r => r[0] !== '')
    .map(r => ({
      id:     String(r[0]),
      label:  r[1],
      icon:   r[2],
      color:  r[3],
      budget: Number(r[4]),
    }));
}

// ── Escribir ───────────────────────────────────────────

function syncExpenses(expenses) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!sheet) sheet = ss.insertSheet(SHEET_EXPENSES);

  sheet.clearContents();
  const headers = ['ID', 'Fecha', 'Descripción', 'Categoría', 'Monto'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  expenses.forEach(exp => {
    sheet.appendRow([exp.id, exp.date, exp.description, exp.category, exp.amount]);
  });

  if (expenses.length > 0) {
    sheet.getRange(2, 5, expenses.length, 1).setNumberFormat('$#,##0.00');
  }
  sheet.autoResizeColumns(1, headers.length);
}

function syncCategories(categories) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CATEGORIES);
  if (!sheet) sheet = ss.insertSheet(SHEET_CATEGORIES);

  sheet.clearContents();
  const headers = ['ID', 'Nombre', 'Ícono', 'Color', 'Presupuesto Mensual'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  categories.forEach(cat => {
    sheet.appendRow([cat.id, cat.label, cat.icon, cat.color, cat.budget || 0]);
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
