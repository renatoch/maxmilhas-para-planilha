var PAYMENT_INFORMED_COLUMN = 11;
var TOTAL_PAYMENT_INFORMED_COLUMN = 12;
var DATE_PAYMENT_INFORMED_COLUMN = 13;

function updatePaymentInfo() {
    var spreadSheetApp, sheet, gmailThreads, firstMessageSubject, firstMessage, firstMessagePlainBody, i,
      method, date, eTicket, paymentValue, totalPayment;
  //, account, airline, airmilesAmount, saleValue, saleValuePerMile, estimatedReceiveDate, boardingFee, luggageFee;

  method = "MaxMilhas"
  
  //spreadSheetApp = SpreadsheetApp.openById('1vUk2OVjnUDuS9mHpF98ILnI1kELnhxl7p16SCxgI9LI')
  spreadSheetApp = SpreadsheetApp.getActiveSpreadsheet();
  sheet = spreadSheetApp.getActiveSheet();
  //gmailThreads=GmailApp.getInboxThreads();
  
  gmailThreads = GmailApp.getUserLabelByName("Venda de milhas").getThreads();
  
  gmailThreads.reverse();
  
  var eticketsMapping = getEticketsMapping(spreadSheetApp);
  
  for (i in gmailThreads) {
    
    firstMessageSubject = gmailThreads[i].getFirstMessageSubject();

    if (firstMessageSubject.search("Seu pagamento foi realizado")>-1 && firstMessageSubject.search("Re:")==-1) {
      firstMessage = gmailThreads[i].getMessages()[0];
      firstMessagePlainBody = firstMessage.getPlainBody().replace(/\r?\n|\r/g," ").replace(/\*/g,"");
      date = firstMessage.getDate();

      var regExp = new RegExp("e-ticket: (\\w*)\\s*Voo: \\d+ Milhas: [\\d\\.]* \\(\\w*\\) Data de emissão: \\d+/\\d+/\\d+ Valor: R\\$ ([\\d,]*)", "gi"); 
      var data;
      
      totalPayment = getTotalPayment(firstMessagePlainBody);
      
      while ((data = regExp.exec(firstMessagePlainBody)) !== null) { 
        eTicket = data[1];
        paymentValue =  data[2];
        
        if(eTicket in eticketsMapping) {
          var row = eticketsMapping[eTicket];
          var isFilled = sheet.getRange(row, PAYMENT_INFORMED_COLUMN).getDisplayValue().trim() != "";
          if (!isFilled) {
            sheet.getRange(row, PAYMENT_INFORMED_COLUMN).setValue(paymentValue);
            sheet.getRange(row, PAYMENT_INFORMED_COLUMN).setNumberFormat("R$ #,##0.00;R$ (#,##0.00)");
            sheet.getRange(row, TOTAL_PAYMENT_INFORMED_COLUMN).setValue(totalPayment);
            sheet.getRange(row, TOTAL_PAYMENT_INFORMED_COLUMN).setNumberFormat("R$ #,##0.00;R$ (#,##0.00)");
            sheet.getRange(row, DATE_PAYMENT_INFORMED_COLUMN).setValue(date);
          }
        }
      }
    }
  }
}

function getTotalPayment(messageBody){
  return getSubstringInTheMiddle(messageBody, "Valor da transferência total: ", ".   Dependendo de seu")
}
