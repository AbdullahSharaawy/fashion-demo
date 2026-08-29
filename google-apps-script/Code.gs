function doPost(e) {
  var order = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Orders");
    sheet.appendRow([
      "رقم الطلب",
      "التاريخ",
      "اسم العميل",
      "الهاتف",
      "نوع الطلب",
      "العنوان",
      "الأصناف",
      "الإجمالي", 
      "طريقة الدفع",
      "الإيصال",
      "ملاحظات",
      "الحالة"
    ]);
  }

  var items = (order.items || []).map(function(item) {
    return item.qty + "x " + item.name + " (" + item.price + ")";
  }).join("؛ ");

  sheet.appendRow([
    order.id || "",
    order.timestamp || new Date().toISOString(),
    order.customerName || "",
    order.phone || "",
    order.orderType || "",
    order.address || "",
    items,
    order.total || 0,
    order.paymentMethod || "",
    order.receiptAttached ? "نعم" : "لا",
    order.notes || "",
    order.status || "جديد"
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
