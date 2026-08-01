function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e.parameter.action;
  
  // 0. 管理員登入驗證
  if (action === "adminLogin") {
    var user = e.parameter.user;
    var pass = e.parameter.pass;
    var adminSheet = ss.getSheetByName("管理員資料");
    var rows = adminSheet ? adminSheet.getDataRange().getValues() : [];
    
    var isAuthenticated = false;
    // 從第 2 列開始比對 (略過標題)
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === user && rows[i][1].toString() === pass) {
        isAuthenticated = true;
        break;
      }
    }
    
    var result = isAuthenticated ? { status: "success" } : { status: "fail" };
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 1. 後台取得所有訂單與客戶資料
  if (action === "getAdminData") {
    var orderSheet = ss.getSheetByName("訂單");
    var customerSheet = ss.getSheetByName("客戶資料");
    
    var orders = orderSheet ? orderSheet.getDataRange().getValues() : [];
    var customers = customerSheet ? customerSheet.getDataRange().getValues() : [];
    
    return ContentService
      .createTextOutput(JSON.stringify({ orders: orders, customers: customers }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. 前台透過電話號碼查詢訂單進度
  if (action === "searchOrder") {
    var phone = e.parameter.phone;
    var orderSheet = ss.getSheetByName("訂單");
    var rows = orderSheet ? orderSheet.getDataRange().getValues() : [];
    var userOrders = [];
    
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][2].toString().trim() === phone.trim()) {
        userOrders.push({
          orderId: rows[i][0],
          name: rows[i][1],
          phone: rows[i][2],
          product: rows[i][3],
          can: rows[i][4],
          box: rows[i][5],
          total: rows[i][6],
          status: rows[i][7],
          trackingCode: rows[i][8] || "無",
          date: rows[i][9]
        });
      }
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", orders: userOrders }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 3. 原本的商品資料讀取
  var sheet = ss.getSheetByName("商品資料");
  var productId = e.parameter.id;
  var rows = sheet.getDataRange().getValues();
  
  var productData = { name: productId, description: "找不到介紹", teaPrice: 0, cans: [], boxes: [] };
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === productId) {
      productData.description = rows[i][1];
      productData.teaPrice = rows[i][2];
      
      var canNames = rows[i][3] ? rows[i][3].toString().split(",") : [];
      var canPrices = rows[i][4] ? rows[i][4].toString().split(",") : [];
      for (var j = 0; j < canNames.length; j++) {
        if(canNames[j].trim() !== "") productData.cans.push({ name: canNames[j].trim(), price: canPrices[j] ? canPrices[j].trim() : 0 });
      }
      
      var boxNames = rows[i][5] ? rows[i][5].toString().split(",") : [];
      var boxPrices = rows[i][6] ? rows[i][6].toString().split(",") : [];
      for (var k = 0; k < boxNames.length; k++) {
        if(boxNames[k].trim() !== "") productData.boxes.push({ name: boxNames[k].trim(), price: boxPrices[k] ? boxPrices[k].trim() : 0 });
      }
      break;
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(productData)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  
  if (data.action === "updateStatus") {
    var orderSheet = ss.getSheetByName("訂單");
    var rows = orderSheet.getDataRange().getValues();
    
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.orderId.toString()) {
        orderSheet.getRange(i + 1, 8).setValue(data.status);
        orderSheet.getRange(i + 1, 9).setValue(data.trackingCode);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  }
  
  var customerSheet = ss.getSheetByName("客戶資料");
  if (!customerSheet) {
    customerSheet = ss.insertSheet("客戶資料");
    customerSheet.appendRow(["姓名", "電話", "住址", "建立時間"]);
  }
  customerSheet.appendRow([data.name, data.phone, data.address, new Date()]);
  
  var orderSheet = ss.getSheetByName("訂單");
  if (!orderSheet) {
    orderSheet = ss.insertSheet("訂單");
    orderSheet.appendRow(["訂單編號", "姓名", "電話", "商品", "罐子", "禮盒", "總金額", "處理進度", "郵局寄件代號", "下單時間"]);
  }
  
  var orderId = "ORD" + new Date().getTime();
  orderSheet.appendRow([
    orderId, data.name, data.phone, data.product, data.can, data.box, data.total,
    "備貨處理中", "無", new Date()
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", orderId: orderId })).setMimeType(ContentService.MimeType.JSON);
}