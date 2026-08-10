// ===== SETTINGS — change this to your own secret word =====
var ADMIN_KEY = "sm-owner-2026"; // used to view and manage orders — keep private

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  // ---- Admin: Update Order Status ----
  if (e.parameter.action === 'updateStatus') {
    if (e.parameter.key !== ADMIN_KEY) {
      return jsonOut({ success: false, error: 'unauthorized' });
    }
    var targetOrderId = (e.parameter.orderId || '').trim().toUpperCase();
    var newStatus = (e.parameter.status || '').trim();
    if (!targetOrderId || !newStatus) {
      return jsonOut({ success: false, error: 'missing orderId or status' });
    }

    var statusColIdx = headers.indexOf('Status');
    if (statusColIdx === -1) statusColIdx = 4; // fallback column E

    for (var k = 1; k < data.length; k++) {
      if (String(data[k][0]).trim().toUpperCase() === targetOrderId) {
        sheet.getRange(k + 1, statusColIdx + 1).setValue(newStatus);
        return jsonOut({ success: true, orderId: targetOrderId, newStatus: newStatus });
      }
    }
    return jsonOut({ success: false, error: 'order not found' });
  }

  // ---- Admin: List orders (optionally filtered to today) ----
  if (e.parameter.action === 'list') {
    if (e.parameter.key !== ADMIN_KEY) {
      return jsonOut({ error: 'unauthorized' });
    }
    var onlyToday = e.parameter.today === '1';
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    var orders = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      headers.forEach(function (h, idx) { obj[h] = row[idx]; });
      if (!onlyToday || String(obj.OrderDate).indexOf(todayStr) !== -1) {
        orders.push(obj);
      }
    }
    return jsonOut({ orders: orders.reverse() });
  }

  // ---- User: List all orders for a specific phone number ("Your Orders") ----
  if (e.parameter.action === 'userOrders') {
    var userPhone = (e.parameter.phone || '').replace(/\D/g, '').slice(-10);
    if (!userPhone) {
      return jsonOut({ success: false, error: 'Phone number required' });
    }

    var userOrderList = [];
    for (var m = 1; m < data.length; m++) {
      var rRow = data[m];
      var rowPhoneNum = String(rRow[1]).replace(/\D/g, '').slice(-10);

      if (rowPhoneNum === userPhone) {
        var orderObj = {};
        headers.forEach(function (h, idx) { orderObj[h] = rRow[idx]; });
        userOrderList.push(orderObj);
      }
    }
    return jsonOut({ success: true, orders: userOrderList.reverse() });
  }

  // ---- Customer: look up a single order by Order ID & Phone ----
  var orderId = (e.parameter.orderId || '').trim().toUpperCase();
  var phone = (e.parameter.phone || '').replace(/\D/g, '');
  var phoneLast10 = phone.slice(-10);

  for (var j = 1; j < data.length; j++) {
    var r = data[j];
    var rowOrderId = String(r[0]).trim().toUpperCase();
    var rowPhone = String(r[1]).replace(/\D/g, '').slice(-10);

    if (rowOrderId === orderId && (phoneLast10 === '' || rowPhone === phoneLast10)) {
      var result = {};
      headers.forEach(function (h, idx) { result[h] = r[idx]; });
      return jsonOut({ found: true, order: result });
    }
  }

  return jsonOut({ found: false });
}

// ---- Customer: place a new order from the website ----
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
    var payload = JSON.parse(e.postData.contents);

    var orderId = 'SM' + Math.floor(10000 + Math.random() * 89999);
    var orderDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');

    sheet.appendRow([
      orderId,
      payload.phone || '',
      payload.name || '',
      orderDate,
      'Placed', // initial status: Placed -> Preparing -> Out for Delivery -> Delivered
      payload.trackingId || '',
      payload.items || '',
      payload.total || '',
      payload.address || '',
      payload.paymentMode || 'COD',
      payload.instructions || ''
    ]);

    return jsonOut({ success: true, orderId: orderId });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
