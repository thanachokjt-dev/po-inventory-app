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
  if (!poId) {
    throw new Error('poId is required');
  }

  var poSheet = getPoMasterSheet_();
  var poValues = poSheet.getDataRange().getValues();
  if (poValues.length <= 1) {
    throw new Error('PO not found: ' + poId);
  }

  var poHeaders = poValues[0];
  var poHeaderIndex = buildHeaderIndex_(poHeaders);
  var headerRow = null;
  for (var i = 1; i < poValues.length; i++) {
    if (poValues[i][poHeaderIndex['po_id']] == poId) {
      headerRow = poValues[i];
      break;
    }
  }

  if (!headerRow) {
    throw new Error('PO not found: ' + poId);
  }

  var header = {};
  poHeaders.forEach(function (name, idx) {
    header[name] = headerRow[idx];
  });

  var itemsSheet = getPoItemsSheet_();
  var itemValues = itemsSheet.getDataRange().getValues();
  var itemHeaders = itemValues[0];
  var itemHeaderIndex = buildHeaderIndex_(itemHeaders);
  var items = [];

  for (var j = 1; j < itemValues.length; j++) {
    var row = itemValues[j];
    if (row[itemHeaderIndex['po_id']] == poId) {
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
  var list = [];
  var stats = {};

  for (var i = 1; i < poValues.length; i++) {
    var row = poValues[i];
    var status = row[headerIndex['status_stage']] || 'Unknown';
    stats[status] = (stats[status] || 0) + 1;

    list.push({
      po_id: row[headerIndex['po_id']],
      po_date: row[headerIndex['po_date']],
      supplier_name: row[headerIndex['supplier_name']],
      po_amount_foreign: row[headerIndex['po_amount_foreign']],
      currency: row[headerIndex['currency']],
      status_stage: row[headerIndex['status_stage']],
      eta_date: row[headerIndex['eta_date']],
      wh_received_date: row[headerIndex['wh_received_date']]
    });
  }

  return {
    stats: stats,
    list: list
  };
}
