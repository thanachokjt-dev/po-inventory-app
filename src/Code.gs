// Entry point for the web app. Returns the HTML file.
// Keeping this as a simple global function avoids "Script function not found: doGet"
// errors when deploying as a Web App from the bound spreadsheet.
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template
    .evaluate()
    .setTitle('PO System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetch all PO rows for the simple dashboard table.
 * Reads the "POs" sheet (header row at row 1) and maps to plain objects
 * so the client can render them directly. Columns are assumed to be:
 * PO ID | Date | Supplier | Amount | Status | ETA | WH Received
 */
function getAllPOsForDashboard() {
  var sheet = getSheet_('POs');
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }

  var dataRows = values.slice(1);
  return dataRows.map(function (row) {
    return {
      poId: row[0],
      date: row[1],
      supplier: row[2],
      amount: row[3],
      status: row[4],
      eta: row[5],
      whReceived: row[6]
    };
  });
}

// Helper to get a sheet by name from the active spreadsheet
function getSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet not found: ' + name);
  }
  return sheet;
}

function getProductSheet_() {
  return getSheet_('PRODUCT_MASTER');
}

function getSupplierSheet_() {
  return getSheet_('SUPPLIER_MASTER');
}

function getPoMasterSheet_() {
  return getSheet_('PO_MASTER');
}

function getPoItemsSheet_() {
  return getSheet_('PO_ITEMS');
}

function getPoHistorySheet_() {
  return getSheet_('PO_HISTORY');
}

// Utility to build a map of header names to column indexes
function buildHeaderIndex_(headers) {
  var map = {};
  headers.forEach(function (name, idx) {
    map[name] = idx;
  });
  return map;
}

/**
 * Normalize a PO id for comparison: convert to string and trim spaces.
 */
function normalizePoId_(value) {
  return value == null ? '' : String(value).trim();
}

/**
 * Resolve a PO id column index, tolerant to different header spellings.
 * Supports variations like "po_id", "PO ID", "Po ID", "PO_ID".
 */
function resolvePoIdColumnIndex_(headers) {
  var normalized = headers.map(function (h) {
    return String(h || '').trim().toLowerCase();
  });
  var candidates = ['po_id', 'po id', 'poid'];
  for (var i = 0; i < normalized.length; i++) {
    if (candidates.indexOf(normalized[i]) !== -1) {
      return i;
    }
  }
  // Fallback to the first column when nothing matches so the lookup still works.
  return 0;
}

// --------------------- Product lookup ---------------------
function listProductsForUi() {
  var sheet = getProductSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }
  var headers = values[0];
  var headerIndex = buildHeaderIndex_(headers);
  var rows = values.slice(1);
  return rows.map(function (row) {
    return {
      sku: row[headerIndex['SKU']],
      productName: row[headerIndex['Product Name']],
      fullName: row[headerIndex['Full Name']],
      color: row[headerIndex['Color']],
      size: row[headerIndex['Size']],
      productImage: row[headerIndex['Product Image']],
      variantImage: row[headerIndex['Variant Image']]
    };
  });
}

function getProductBySku(sku) {
  if (!sku) return null;
  var sheet = getProductSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return null;
  }
  var headers = values[0];
  var headerIndex = buildHeaderIndex_(headers);
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[headerIndex['SKU']] == sku) {
      return {
        sku: row[headerIndex['SKU']],
        productName: row[headerIndex['Product Name']],
        fullName: row[headerIndex['Full Name']],
        color: row[headerIndex['Color']],
        size: row[headerIndex['Size']],
        productImage: row[headerIndex['Product Image']],
        variantImage: row[headerIndex['Variant Image']]
      };
    }
  }
  return null;
}

// --------------------- Supplier lookup ---------------------
function listSuppliers() {
  var sheet = getSupplierSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }
  var headers = values[0];
  var headerIndex = buildHeaderIndex_(headers);
  var rows = values.slice(1);
  return rows.map(function (row) {
    return {
      supplier_code: row[headerIndex['supplier_code']],
      supplier_name: row[headerIndex['supplier_name']],
      payment_terms_text: row[headerIndex['payment_terms_text']],
      currency: row[headerIndex['currency']],
      incoterm: row[headerIndex['incoterm']],
      ship_mode: row[headerIndex['ship_mode']],
      contact_name: row[headerIndex['contact_name']],
      contact_email: row[headerIndex['contact_email']],
      bank_detail: row[headerIndex['bank_detail']]
    };
  });
}

function getSupplierByCode(code) {
  if (!code) return null;
  var sheet = getSupplierSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return null;
  }
  var headers = values[0];
  var headerIndex = buildHeaderIndex_(headers);
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[headerIndex['supplier_code']] == code) {
      return {
        supplier_code: row[headerIndex['supplier_code']],
        supplier_name: row[headerIndex['supplier_name']],
        payment_terms_text: row[headerIndex['payment_terms_text']],
        currency: row[headerIndex['currency']],
        incoterm: row[headerIndex['incoterm']],
        ship_mode: row[headerIndex['ship_mode']],
        contact_name: row[headerIndex['contact_name']],
        contact_email: row[headerIndex['contact_email']],
        bank_detail: row[headerIndex['bank_detail']]
      };
    }
  }
  return null;
}

// --------------------- Save PO with items ---------------------
function savePoWithItems(payload) {
  if (!payload || !payload.header || !payload.header.po_id) {
    throw new Error('Missing PO header or po_id');
  }
  var header = payload.header;
  var items = payload.items || [];
  var poId = header.po_id;

  // Save or update PO_MASTER
  var poSheet = getPoMasterSheet_();
  var poValues = poSheet.getDataRange().getValues();
  var poHeaders = poValues[0];
  var poHeaderIndex = buildHeaderIndex_(poHeaders);
  var poRowNumber = null;
  for (var i = 1; i < poValues.length; i++) {
    if (poValues[i][poHeaderIndex['po_id']] == poId) {
      poRowNumber = i + 1; // Convert to 1-based row number
      break;
    }
  }
  var poRowData = poHeaders.map(function (name) {
    return header[name] !== undefined ? header[name] : '';
  });
  if (poRowNumber) {
    poSheet.getRange(poRowNumber, 1, 1, poHeaders.length).setValues([poRowData]);
  } else {
    poSheet.appendRow(poRowData);
  }

  // Delete existing items for this PO
  var itemsSheet = getPoItemsSheet_();
  var lastRow = itemsSheet.getLastRow();
  for (var row = lastRow; row >= 2; row--) {
    var value = itemsSheet.getRange(row, 1).getValue();
    if (value == poId) {
      itemsSheet.deleteRow(row);
    }
  }

  // Append new items
  if (items.length > 0) {
    var itemRows = items.map(function (item) {
      return [
        poId,
        item.line_no,
        item.sku,
        item.product_title,
        item.variant_title,
        item.image_url,
        item.qty,
        item.unit_price,
        item.line_amount,
        item.currency,
        item.remark
      ];
    });
    itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
  }

  // Optional snapshot to PO_HISTORY
  if (header.status_stage == 'In warehouse' || header.status_stage == 'Closed') {
    var historySheet = getPoHistorySheet_();
    var snapshot = [
      poId,
      new Date(),
      header.status_stage,
      header.po_amount_foreign,
      header.po_amount_thb,
      header.supplier_code,
      header.supplier_name,
      header.eta_date,
      header.wh_received_date,
      header.remark
    ];
    historySheet.appendRow(snapshot);
  }
}

// --------------------- Load PO with items ---------------------
function getPoWithItems(poId) {
  var targetId = normalizePoId_(poId);
  if (!targetId) {
    throw new Error('poId is required');
  }

  var poSheet = getPoMasterSheet_();
  var poValues = poSheet.getDataRange().getValues();
  if (poValues.length <= 1) {
    throw new Error('PO not found: ' + targetId);
  }

  var poHeaders = poValues[0];
  var poIdColIdx = resolvePoIdColumnIndex_(poHeaders);
  var poHeaderIndex = buildHeaderIndex_(poHeaders);
  var headerRow = null;

  for (var i = 1; i < poValues.length; i++) {
    var row = poValues[i];
    if (normalizePoId_(row[poIdColIdx]) === targetId) {
      headerRow = row;
      break;
    }
  }

  if (!headerRow) {
    throw new Error('PO not found: ' + targetId);
  }

  // Map every header name to its value so the client receives the full row.
  var header = {};
  poHeaders.forEach(function (name, idx) {
    header[name] = headerRow[idx];
  });

  var itemsSheet = getPoItemsSheet_();
  var itemValues = itemsSheet.getDataRange().getValues();
  var items = [];

  if (itemValues.length > 1) {
    var itemHeaders = itemValues[0];
    var itemPoIdColIdx = resolvePoIdColumnIndex_(itemHeaders);
    var itemHeaderIndex = buildHeaderIndex_(itemHeaders);

    for (var j = 1; j < itemValues.length; j++) {
      var row = itemValues[j];
      if (normalizePoId_(row[itemPoIdColIdx]) === targetId) {
        items.push({
          line_no: row[itemHeaderIndex['line_no']],
          sku: row[itemHeaderIndex['sku']],
          product_title: row[itemHeaderIndex['product_title']],
          variant_title: row[itemHeaderIndex['variant_title']],
          image_url: row[itemHeaderIndex['image_url']],
          qty: row[itemHeaderIndex['qty']],
          unit_price: row[itemHeaderIndex['unit_price']],
          line_amount: row[itemHeaderIndex['line_amount']],
          currency: row[itemHeaderIndex['currency']],
          remark: row[itemHeaderIndex['remark']]
        });
      }
    }
  }

  items.sort(function (a, b) {
    return Number(a.line_no) - Number(b.line_no);
  });

  return {
    header: header,
    items: items
  };
}

// --------------------- Dashboard data ---------------------
function getPoDashboardData() {
  var poSheet = getPoMasterSheet_();
  var poValues = poSheet.getDataRange().getValues();
  if (poValues.length <= 1) {
    return { stats: {}, list: [] };
  }

  var headers = poValues[0];
  var headerIndex = buildHeaderIndex_(headers);
  var poIdColIdx = resolvePoIdColumnIndex_(headers);

  var list = [];
  var stats = {};
  for (var i = 1; i < poValues.length; i++) {
    var row = poValues[i];
    var poId = normalizePoId_(row[poIdColIdx]);
    if (!poId) {
      continue; // skip blank IDs so dashboard only shows real POs
    }

    // Pull values defensively in case a column is missing from this sheet
    var statusIdx = headerIndex['status_stage'];
    var supplierIdx = headerIndex['supplier_name'];
    var amountIdx = headerIndex['po_amount_foreign'];
    var currencyIdx = headerIndex['currency'];
    var poDateIdx = headerIndex['po_date'];
    var etaIdx = headerIndex['eta_date'];
    var whIdx = headerIndex['wh_received_date'];

    var status = statusIdx == null ? '' : row[statusIdx];
    var statusLabel = status ? status : 'Unknown';
    stats[statusLabel] = (stats[statusLabel] || 0) + 1;

    list.push({
      po_id: poId,
      po_date: poDateIdx == null ? '' : row[poDateIdx],
      supplier_name: supplierIdx == null ? '' : row[supplierIdx],
      po_amount_foreign: amountIdx == null ? '' : row[amountIdx],
      currency: currencyIdx == null ? '' : row[currencyIdx],
      status_stage: statusLabel,
      eta_date: etaIdx == null ? '' : row[etaIdx],
      wh_received_date: whIdx == null ? '' : row[whIdx]
    });
  }

  return {
    stats: stats,
    list: list
  };
}
