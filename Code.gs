// ============ GROCERY POS - GOOGLE APPS SCRIPT ============
const SS_ID_KEY = 'SPREADSHEET_ID';
const SHEETS = {
  SETTINGS: 'Settings',
  VENDORS: 'Vendors',
  CATEGORIES: 'Categories',
  PRODUCTS: 'Products',
  SALES: 'Sales',
  SALE_ITEMS: 'SaleItems'
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Grocery POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(SS_ID_KEY);
  var ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Grocery POS Database');
    props.setProperty(SS_ID_KEY, ss.getId());
    var def = ss.getSheetByName('Sheet1');
    if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch (e) {} }
  }
  return ss;
}

function getSheet(name) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function initialize() {
  try {
    var ss = getSpreadsheet();
    
    var s = getSheet(SHEETS.SETTINGS);
    if (s.getLastRow() < 2) {
      s.clear();
      s.appendRow(['key', 'value']);
      s.getRange(2, 1, 11, 2).setValues([
        ['store_name', 'GROCERY POS'],
        ['store_address', '123 Market Street, City'],
        ['store_phone', '555-0100'],
        ['currency_symbol', 'RS-'],
        ['theme', 'dark'],
        ['language', 'en'],
        ['admin_password', ''],
        ['receipt_footer', 'Thank you for your purchase!'],
        ['app_footer', 'Powered by Grocery POS v2.0'],
        ['enabled_categories', JSON.stringify(['Fruits','Vegetables','Dairy','Bakery','Beverages','Snacks','Grains'])],
        ['db_mode', 'local']
      ]);
    }
    
    var v = getSheet(SHEETS.VENDORS);
    if (v.getLastRow() < 2) {
      v.clear();
      var now = new Date().toISOString();
      v.appendRow(['id', 'name', 'contact', 'email', 'address', 'created_at']);
      v.getRange(2, 1, 4, 6).setValues([
        [1, 'Fresh Farms Ltd', '555-0101', 'sales@freshfarms.com', '12 Market St', now],
        [2, 'Dairy Delight Co', '555-0102', 'info@dairydelight.com', '88 Dairy Ave', now],
        [3, 'Grain Masters', '555-0103', 'orders@grainmasters.com', '45 Wheat Rd', now],
        [4, 'Beverage World', '555-0104', 'contact@bevworld.com', '99 Drink Blvd', now]
      ]);
    }
    
    var c = getSheet(SHEETS.CATEGORIES);
    if (c.getLastRow() < 2) {
      c.clear();
      c.appendRow(['id', 'name', 'enabled']);
      var cats = ['Fruits', 'Vegetables', 'Dairy', 'Bakery', 'Beverages', 'Snacks', 'Grains'];
      c.getRange(2, 1, cats.length, 3).setValues(cats.map(function(n, i) { return [i + 1, n, true]; }));
    }
    
    var p = getSheet(SHEETS.PRODUCTS);
    if (p.getLastRow() < 2) {
      p.clear();
      var now2 = new Date().toISOString();
      p.appendRow(['id', 'barcode', 'name', 'category_id', 'vendor_id', 'cost_price', 'selling_price', 'stock', 'unit', 'low_stock_threshold', 'created_at']);
      var prods = [
        [1, '8901234567890', 'Fresh Apples 1kg', 1, 1, 2.5, 4.0, 50, 'kg', 10, now2],
        [2, '8901234567891', 'Bananas 1 dozen', 1, 1, 1.8, 3.0, 40, 'dz', 10, now2],
        [3, '8901234567892', 'Tomatoes 1kg', 2, 1, 1.5, 2.5, 60, 'kg', 10, now2],
        [4, '8901234567893', 'Onions 1kg', 2, 1, 1.0, 1.8, 80, 'kg', 10, now2],
        [5, '8901234567894', 'Whole Milk 1L', 3, 2, 1.2, 2.0, 100, 'L', 10, now2],
        [6, '8901234567895', 'Cheddar Cheese 250g', 3, 2, 3.5, 5.5, 30, 'pcs', 10, now2],
        [7, '8901234567896', 'White Bread', 4, 3, 1.0, 1.8, 25, 'pcs', 10, now2],
        [8, '8901234567897', 'Croissants 4pk', 4, 3, 2.5, 4.0, 20, 'pk', 10, now2],
        [9, '8901234567898', 'Cola 1.5L', 5, 4, 1.2, 2.0, 70, 'btl', 10, now2],
        [10, '8901234567899', 'Orange Juice 1L', 5, 4, 2.0, 3.5, 45, 'L', 10, now2],
        [11, '8901234567900', 'Potato Chips 150g', 6, 4, 1.5, 2.5, 55, 'pcs', 10, now2],
        [12, '8901234567901', 'Basmati Rice 5kg', 7, 3, 8.0, 12.0, 35, 'kg', 10, now2],
        [13, '8901234567902', 'Whole Wheat Flour 2kg', 7, 3, 2.5, 4.0, 40, 'kg', 10, now2],
        [14, '8901234567903', 'Yogurt 500g', 3, 2, 1.8, 3.0, 8, 'pcs', 10, now2],
        [15, '8901234567904', 'Eggs 12pk', 3, 2, 2.5, 4.5, 6, 'pk', 10, now2]
      ];
      p.getRange(2, 1, prods.length, 11).setValues(prods);
    }
    
    var sl = getSheet(SHEETS.SALES);
    if (sl.getLastRow() < 1) {
      sl.clear();
      sl.appendRow(['id', 'invoice_no', 'customer_name', 'total', 'discount', 'tax', 'grand_total', 'payment_method', 'amount_paid', 'change_amount', 'created_at']);
    }
    
    var si = getSheet(SHEETS.SALE_ITEMS);
    if (si.getLastRow() < 1) {
      si.clear();
      si.appendRow(['id', 'sale_id', 'product_id', 'product_name', 'quantity', 'price', 'total']);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: 'Database initialized!', url: ss.getUrl() };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function getSettings() {
  try {
    var sheet = getSheet(SHEETS.SETTINGS);
    var data = sheet.getDataRange().getValues();
    var obj = {};
    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      var value = data[i][1];
      if (!key) continue;
      if (key === 'admin_password') {
        obj.has_password = value && String(value).length > 0;
        obj.admin_password = value;
      } else if (key === 'enabled_categories') {
        try { obj.enabled_categories = JSON.parse(value); } catch (e) { obj.enabled_categories = []; }
      } else {
        obj[key] = value;
      }
    }
    return obj;
  } catch (e) {
    return {
      store_name: 'GROCERY POS', currency_symbol: 'RS-', theme: 'dark', language: 'en',
      has_password: false, enabled_categories: ['Fruits','Vegetables','Dairy','Bakery','Beverages','Snacks','Grains'],
      receipt_footer: 'Thank you!', app_footer: 'Powered by Grocery POS'
    };
  }
}

function saveSettings(settings) {
  try {
    var sheet = getSheet(SHEETS.SETTINGS);
    var data = sheet.getDataRange().getValues();
    var keys = {};
    for (var i = 1; i < data.length; i++) { if (data[i][0]) keys[data[i][0]] = i + 1; }
    for (var k in settings) {
      var v = settings[k];
      var value = (k === 'enabled_categories' && Array.isArray(v)) ? JSON.stringify(v) : String(v);
      if (keys[k]) sheet.getRange(keys[k], 2).setValue(value);
      else sheet.appendRow([k, value]);
    }
    return { ok: true };
  } catch (e) { throw new Error(e.message); }
}

function getVendors() {
  try {
    var sheet = getSheet(SHEETS.VENDORS);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    return data.slice(1).filter(function(r) { return r[0]; }).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });
  } catch (e) { return []; }
}

function addVendor(v) {
  var sheet = getSheet(SHEETS.VENDORS);
  var id = getNextId(sheet);
  sheet.appendRow([id, v.name, v.contact || '', v.email || '', v.address || '', new Date().toISOString()]);
  return { id: id };
}

function updateVendor(id, v) {
  var sheet = getSheet(SHEETS.VENDORS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2, 1, 4).setValues([[v.name, v.contact || '', v.email || '', v.address || '']]);
      return { ok: true };
    }
  }
  throw new Error('Vendor not found');
}

function deleteVendor(id) {
  deleteRowById(SHEETS.VENDORS, id);
  return { ok: true };
}

function getCategories() {
  try {
    var sheet = getSheet(SHEETS.CATEGORIES);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    return data.slice(1).filter(function(r) { return r[0]; }).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });
  } catch (e) { return []; }
}

function addCategory(name) {
  var sheet = getSheet(SHEETS.CATEGORIES);
  var id = getNextId(sheet);
  sheet.appendRow([id, name, true]);
  return { id: id };
}

function deleteCategory(id) {
  deleteRowById(SHEETS.CATEGORIES, id);
  return { ok: true };
}

function toggleCategory(id, enabled) {
  var sheet = getSheet(SHEETS.CATEGORIES);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 3).setValue(enabled);
      return { ok: true };
    }
  }
  throw new Error('Category not found');
}

function resetCategories() {
  var sheet = getSheet(SHEETS.CATEGORIES);
  sheet.clear();
  sheet.appendRow(['id', 'name', 'enabled']);
  var cats = ['Fruits', 'Vegetables', 'Dairy', 'Bakery', 'Beverages', 'Snacks', 'Grains'];
  sheet.getRange(2, 1, cats.length, 3).setValues(cats.map(function(c, i) { return [i + 1, c, true]; }));
  return { ok: true };
}

function getProducts() {
  try {
    var sheet = getSheet(SHEETS.PRODUCTS);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var categories = getCategories();
    var vendors = getVendors();
    return data.slice(1).filter(function(r) { return r[0]; }).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      var cat = categories.find(function(c) { return c.id == obj.category_id; });
      var ven = vendors.find(function(v) { return v.id == obj.vendor_id; });
      obj.category_name = cat ? cat.name : '';
      obj.vendor_name = ven ? ven.name : '';
      obj.vendor_contact = ven ? ven.contact : '';
      return obj;
    });
  } catch (e) { return []; }
}

function getLowStock() {
  try {
    return getProducts().filter(function(p) { return p.stock <= p.low_stock_threshold; }).sort(function(a, b) { return a.stock - b.stock; });
  } catch (e) { return []; }
}

function getProductByBarcode(barcode) {
  try {
    if (!barcode || barcode === 'undefined') return null;
    var products = getProducts();
    return products.find(function(p) { return String(p.barcode) === String(barcode); }) || null;
  } catch (e) { return null; }
}

function addProduct(product) {
  var sheet = getSheet(SHEETS.PRODUCTS);
  if (product.custom_category && product.custom_category.trim()) {
    var existing = getCategories().find(function(c) { return c.name.toLowerCase() === product.custom_category.toLowerCase(); });
    if (existing) product.category_id = existing.id;
    else product.category_id = addCategory(product.custom_category.trim()).id;
  }
  var id = getNextId(sheet);
  sheet.appendRow([
    id, product.barcode || '', product.name, product.category_id || '', product.vendor_id || '',
    Number(product.cost_price) || 0, Number(product.selling_price), Number(product.stock) || 0,
    product.unit || 'pcs', Number(product.low_stock_threshold) || 10, new Date().toISOString()
  ]);
  return { id: id };
}

function updateProduct(id, product) {
  var sheet = getSheet(SHEETS.PRODUCTS);
  var data = sheet.getDataRange().getValues();
  if (product.custom_category && product.custom_category.trim()) {
    var existing = getCategories().find(function(c) { return c.name.toLowerCase() === product.custom_category.toLowerCase(); });
    if (existing) product.category_id = existing.id;
    else product.category_id = addCategory(product.custom_category.trim()).id;
  }
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2, 1, 9).setValues([[
        product.barcode || '', product.name, product.category_id || '', product.vendor_id || '',
        Number(product.cost_price) || 0, Number(product.selling_price), Number(product.stock) || 0,
        product.unit || 'pcs', Number(product.low_stock_threshold) || 10
      ]]);
      return { ok: true };
    }
  }
  throw new Error('Product not found');
}

function deleteProduct(id) {
  deleteRowById(SHEETS.PRODUCTS, id);
  return { ok: true };
}

function addSale(saleData) {
  var productsSheet = getSheet(SHEETS.PRODUCTS);
  var salesSheet = getSheet(SHEETS.SALES);
  var saleItemsSheet = getSheet(SHEETS.SALE_ITEMS);
  var products = getProducts();
  
  for (var x = 0; x < saleData.items.length; x++) {
    var it = saleData.items[x];
    var p = products.find(function(x) { return x.id == it.product_id; });
    if (!p || p.stock < it.quantity) throw new Error('Insufficient stock: ' + (p ? p.name : it.product_id));
  }
  
  var subtotal = 0;
  for (var x = 0; x < saleData.items.length; x++) subtotal += saleData.items[x].price * saleData.items[x].quantity;
  var tax = Number(((subtotal * (saleData.taxRate || 0)) / 100).toFixed(2));
  var grand_total = Number((subtotal - (saleData.discount || 0) + tax).toFixed(2));
  var change = Number(((saleData.amount_paid || 0) - grand_total).toFixed(2));
  var d = new Date();
  var invoice_no = 'INV-' + d.getFullYear().toString().slice(-2) + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0') + '-' + Math.floor(Math.random() * 9000 + 1000);
  
  var saleId = getNextId(salesSheet);
  salesSheet.appendRow([
    saleId, invoice_no, saleData.customer_name || 'Walking Customer', subtotal, saleData.discount || 0,
    tax, grand_total, saleData.payment_method || 'cash', saleData.amount_paid || 0, change, new Date().toISOString()
  ]);
  
  var productsData = productsSheet.getDataRange().getValues();
  var itemsToReturn = [];
  
  for (var x = 0; x < saleData.items.length; x++) {
    var it = saleData.items[x];
    var itemId = getNextId(saleItemsSheet);
    var p = products.find(function(x) { return x.id == it.product_id; });
    saleItemsSheet.appendRow([itemId, saleId, it.product_id, p ? p.name : 'Item', it.quantity, it.price, Number((it.price * it.quantity).toFixed(2))]);
    itemsToReturn.push({ product_id: it.product_id, name: p ? p.name : 'Item', quantity: it.quantity, price: it.price, total: Number((it.price * it.quantity).toFixed(2)) });
    
    for (var i = 1; i < productsData.length; i++) {
      if (String(productsData[i][0]) === String(it.product_id)) {
        productsSheet.getRange(i + 1, 8).setValue(Number(productsData[i][7]) - it.quantity);
        productsData[i][7] = Number(productsData[i][7]) - it.quantity;
        break;
      }
    }
  }
  
  return {
    id: saleId, invoice_no: invoice_no, customer_name: saleData.customer_name || 'Walking Customer',
    subtotal: subtotal, discount: saleData.discount || 0, tax: tax, grand_total: grand_total,
    payment_method: saleData.payment_method || 'cash', amount_paid: saleData.amount_paid || 0, change: change, items: itemsToReturn
  };
}

function getSales(filter) {
  try {
    var sheet = getSheet(SHEETS.SALES);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var sales = data.slice(1).filter(function(r) { return r[0]; }).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      var itemsSheet = getSheet(SHEETS.SALE_ITEMS);
      var itemsData = itemsSheet.getDataRange().getValues();
      obj.item_count = itemsData.filter(function(r) { return String(r[1]) === String(obj.id); }).length;
      return obj;
    });
    
    if (filter && filter.period) {
      var now = new Date();
      sales = sales.filter(function(s) {
        var d = new Date(s.created_at);
        if (filter.period === 'hour') return (now - d) < 3600000;
        if (filter.period === 'day') return d.toDateString() === now.toDateString();
        if (filter.period === 'week') return (now - d) < 7 * 86400000;
        if (filter.period === 'month') return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
        if (filter.period === 'year') return d.getFullYear() === now.getFullYear();
        return true;
      });
    } else if (filter && filter.from && filter.to) {
      var from = new Date(filter.from), to = new Date(filter.to);
      sales = sales.filter(function(s) { var d = new Date(s.created_at); return d >= from && d <= to; });
    }
    return sales.sort(function(a, b) { return new Date(b.created_at) - new Date(a.created_at); }).slice(0, 500);
  } catch (e) { return []; }
}

function getSaleById(id) {
  var sheet = getSheet(SHEETS.SALES);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var sale = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sale = {};
      headers.forEach(function(h, j) { sale[h] = data[i][j]; });
      break;
    }
  }
  if (!sale) throw new Error('Sale not found');
  var itemsSheet = getSheet(SHEETS.SALE_ITEMS);
  var itemsData = itemsSheet.getDataRange().getValues();
  var itemsHeaders = itemsData[0];
  sale.items = itemsData.slice(1).filter(function(r) { return String(r[1]) === String(id); }).map(function(row) {
    var obj = {};
    itemsHeaders.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
  return sale;
}

function deleteSale(id) {
  var salesSheet = getSheet(SHEETS.SALES);
  var saleItemsSheet = getSheet(SHEETS.SALE_ITEMS);
  var productsSheet = getSheet(SHEETS.PRODUCTS);
  var sale = getSaleById(id);
  if (!sale) throw new Error('Sale not found');
  
  var productsData = productsSheet.getDataRange().getValues();
  for (var x = 0; x < sale.items.length; x++) {
    var it = sale.items[x];
    for (var i = 1; i < productsData.length; i++) {
      if (String(productsData[i][0]) === String(it.product_id)) {
        productsSheet.getRange(i + 1, 8).setValue(Number(productsData[i][7]) + it.quantity);
        productsData[i][7] = Number(productsData[i][7]) + it.quantity;
        break;
      }
    }
  }
  
  var itemsData = saleItemsSheet.getDataRange().getValues();
  for (var i = itemsData.length - 1; i >= 1; i--) {
    if (String(itemsData[i][1]) === String(id)) saleItemsSheet.deleteRow(i + 1);
  }
  
  var salesData = salesSheet.getDataRange().getValues();
  for (var i = 1; i < salesData.length; i++) {
    if (String(salesData[i][0]) === String(id)) { salesSheet.deleteRow(i + 1); break; }
  }
  return { ok: true };
}

function getDashboard() {
  try {
    var sales = getSales({});
    var products = getProducts();
    var vendors = getVendors();
    var today = new Date().toDateString();
    var todaySales = sales.filter(function(s) { return new Date(s.created_at).toDateString() === today; });
    var todayTotal = todaySales.reduce(function(sum, s) { return sum + Number(s.grand_total); }, 0);
    
    var last7 = [];
    for (var i = 6; i >= 0; i--) {
      var d = new Date(); d.setDate(d.getDate() - i);
      var dayStr = d.toISOString().slice(0, 10);
      var daySales = sales.filter(function(s) { return new Date(s.created_at).toISOString().slice(0, 10) === dayStr; });
      last7.push({ day: dayStr, total: daySales.reduce(function(sum, s) { return sum + Number(s.grand_total); }, 0), count: daySales.length });
    }
    
    return {
      todaySales: todayTotal, todayCount: todaySales.length,
      totalProducts: products.length, lowStock: products.filter(function(p) { return p.stock <= p.low_stock_threshold; }).length,
      totalVendors: vendors.length, last7: last7, topProducts: [], categories: []
    };
  } catch (e) {
    return { todaySales: 0, todayCount: 0, totalProducts: 0, lowStock: 0, totalVendors: 0, last7: [], topProducts: [], categories: [] };
  }
}

function getNextId(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 1;
  var maxId = 0;
  for (var i = 1; i < data.length; i++) {
    var id = Number(data[i][0]);
    if (id > maxId) maxId = id;
  }
  return maxId + 1;
}

function deleteRowById(sheetName, id) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); return; }
  }
}
