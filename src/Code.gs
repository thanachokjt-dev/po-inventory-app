const PO_SHEET_NAME = 'PO_MASTER';
const ITEM_SHEET_NAME = 'PO_ITEMS';

const PO_HEADERS = [
  'po_id',
  'supplier_id',
  'deposit_pct',
  'conditions_pay',
  'po_owner',
  'allstatus',
  'Formulas_Check',
  'Note',
  'po_sent_date',
  'order_date',
  'approval_date',
  'approval_owners',
  'approval_status',
  'Total_Amount',
  'deposit_amount',
  'status_deposit_r1',
  'deposit_due_date',
  'deposit_paid_date',
  'requested_ship_date',
  'confirmed_ship_date',
  'deposit_amount_R2',
  'status_deposit_R2',
  'deposit_R2_due_date',
  'deposit_R2_paid_date',
  'etd_port',
  'eta_port',
  'requested_delivery_date',
  'delivery_date',
  'status',
  'delivery_reciver',
  'delivery_recived',
  'delivery_note',
  'solution_comment',
  'date_note',
  'clearance_start_date',
  'clearance_complete_date',
  'grn_date',
  'invoice_no',
  'balance_due_date',
  'balance_paid_date',
  'freight_mode',
  'incoterms',
  'awb_bl_no',
  'budget_bucket',
  'notes',
  'status_stage_index',
  'otif_flag',
  'deposit_amount_R3',
  'status_deposit_R3',
  'deposit_R3_due_date',
  'deposit_R3_paid_date',
  'ETA_score',
  'ETA_TAKE',
  'Leadtime_days',
  'ETA_DATA',
  'ETA_SUP_COMMIT',
  'Real_Recived'
];

const ITEM_HEADERS = [
  'po_id',
  'line_no',
  'description',
  'quantity',
  'unit_price',
  'currency',
  'requested_delivery_date',
  'confirmed_delivery_date',
  'received_quantity',
  'receiving_note'
];

function doGet() {
  setupSheets();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PO Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName(PO_SHEET_NAME) || ss.insertSheet(PO_SHEET_NAME);
  const itemSheet = ss.getSheetByName(ITEM_SHEET_NAME) || ss.insertSheet(ITEM_SHEET_NAME);

  if (poSheet.getLastRow() === 0) {
    poSheet.appendRow(PO_HEADERS);
  }

  if (itemSheet.getLastRow() === 0) {
    itemSheet.appendRow(ITEM_HEADERS);
  }
}

function createPo(payload) {
  setupSheets();
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName(PO_SHEET_NAME);
  const itemSheet = ss.getSheetByName(ITEM_SHEET_NAME);

  const po = payload.po || {};
  const items = payload.items || [];

  const poId = po.po_id || `PO-${new Date().getTime()}`;
  po.po_id = poId;

  const depositPct = Number(po.deposit_pct || 0);
  const totalAmount = Number(po.Total_Amount || 0);
  po.deposit_amount = depositPct ? (totalAmount * depositPct) / 100 : po.deposit_amount;

  const row = PO_HEADERS.map(header => po[header] || '');
  poSheet.appendRow(row);

  if (items.length) {
    const itemRows = items.map((item, index) => (
      [
        poId,
        item.line_no || index + 1,
        item.description || '',
        Number(item.quantity || 0),
        Number(item.unit_price || 0),
        item.currency || 'USD',
        item.requested_delivery_date || '',
        item.confirmed_delivery_date || '',
        Number(item.received_quantity || 0),
        item.receiving_note || ''
      ]
    ));
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
  }

  return { po_id: poId };
}

function listPos() {
  setupSheets();
  const sheet = SpreadsheetApp.getActive().getSheetByName(PO_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const [headerRow, ...rows] = values;

  return rows.map(row => headerRow.reduce((acc, key, idx) => {
    acc[key] = row[idx];
    return acc;
  }, {})).filter(record => record.po_id);
}

function getDashboardData() {
  const pos = listPos();

  const summary = pos.reduce((acc, po) => {
    const status = po.status || 'Unspecified';
    acc.counts[status] = (acc.counts[status] || 0) + 1;
    const supplierId = po.supplier_id || 'Unknown';
    acc.suppliers[supplierId] = acc.suppliers[supplierId] || { total: 0, leadTimes: [] };
    acc.suppliers[supplierId].total += 1;

    const depositDate = parseDate(po.deposit_paid_date);
    const deliveryDate = parseDate(po.delivery_date || po.Real_Recived);
    if (depositDate && deliveryDate) {
      const lead = Math.round((deliveryDate - depositDate) / (1000 * 60 * 60 * 24));
      acc.suppliers[supplierId].leadTimes.push(lead);
    }
    return acc;
  }, { counts: {}, suppliers: {} });

  const supplierLeadTimes = Object.keys(summary.suppliers).map(supplierId => {
    const info = summary.suppliers[supplierId];
    const count = info.leadTimes.length;
    const avg = count ? Math.round(info.leadTimes.reduce((a, b) => a + b, 0) / count) : null;
    return { supplier_id: supplierId, pos: info.total, avg_leadtime_days: avg, samples: count };
  });

  return { counts: summary.counts, pos: pos.slice(-20).reverse(), suppliers: supplierLeadTimes };
}

function parseDate(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }
  const asNumber = Number(value);
  if (!isNaN(asNumber)) {
    return new Date(asNumber);
  }
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function updateReceiving(data) {
  setupSheets();
  const sheet = SpreadsheetApp.getActive().getSheetByName(PO_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const [headers, ...rows] = values;
  const poIndex = headers.indexOf('po_id');
  const deliveryDateIdx = headers.indexOf('delivery_date');
  const receiverIdx = headers.indexOf('delivery_reciver');
  const receivedIdx = headers.indexOf('delivery_recived');
  const noteIdx = headers.indexOf('delivery_note');
  const statusIdx = headers.indexOf('status');

  const poRowIndex = rows.findIndex(row => row[poIndex] === data.po_id);
  if (poRowIndex === -1) {
    throw new Error('PO not found');
  }

  const rowNumber = poRowIndex + 2; // account for header row
  if (deliveryDateIdx > -1 && data.delivery_date) {
    sheet.getRange(rowNumber, deliveryDateIdx + 1).setValue(data.delivery_date);
  }
  if (receiverIdx > -1 && data.delivery_reciver) {
    sheet.getRange(rowNumber, receiverIdx + 1).setValue(data.delivery_reciver);
  }
  if (receivedIdx > -1 && data.delivery_recived) {
    sheet.getRange(rowNumber, receivedIdx + 1).setValue(data.delivery_recived);
  }
  if (noteIdx > -1 && data.delivery_note) {
    sheet.getRange(rowNumber, noteIdx + 1).setValue(data.delivery_note);
  }
  if (statusIdx > -1 && data.status) {
    sheet.getRange(rowNumber, statusIdx + 1).setValue(data.status);
  }

  return { po_id: data.po_id };
}
