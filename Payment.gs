var PAYMENT_INFORMED_COLUMN = 11;
var TOTAL_PAYMENT_INFORMED_COLUMN = 12;
var DATE_PAYMENT_INFORMED_COLUMN = 13;

function updatePaymentInfo() {
    var sheet, gmailThreads, subject, message, messages, plainBody, i,
      method, date, eTicket, paymentValue, totalPayment;

  method = "MaxMilhas"
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Import");
  if (sheet == null) {
    throw new Error("There must be an 'Import' tab on the spreadsheet");
  }
  var namedRanges = sheet.getNamedRanges();
  var rangeStatus = getNamedRangeFast(namedRanges, "Status");

  if (rangeStatus != null) {
    rangeStatus.setValue("Getting Gmail threads...");
    rangeStatus.setBackground("#f9cb9c");
  }

  var startDateFilter = getStartDateFilterIfPresent(namedRanges);
  
  gmailThreads = GmailApp.search(startDateFilter + " label:\"Venda de milhas\"");
  gmailThreads.reverse();
  
  if (rangeStatus != null) {
    rangeStatus.setValue("Getting tickets mapping from sheet...");
  }  
  
  var eticketsMapping = getEticketsMapping(sheet);

  if (rangeStatus != null) {
    rangeStatus.setValue("Processing " + gmailThreads.length + " gmail threads...");
    SpreadsheetApp.flush();
  }
  
  for (i in gmailThreads) {
    
    messages = gmailThreads[i].getMessages();
    
    for (var j in messages) {
      message = messages[j];
      subject = message.getSubject();
      
      if (subject.search("Seu pagamento foi realizado")>-1 && subject.search("Re:")==-1) {
        plainBody = removeLineBreaksAndSpecialChars(message.getPlainBody());
        date = message.getDate();
        
        var regExp = new RegExp("e-ticket: (\\w*)\\s*Voo: \\d+ Milhas: [\\d\\.]* \\(\\w*\\) Data de emissão: \\d+/\\d+/\\d+ Valor: R\\$ ([\\d,\\.]*)", "gi"); 
        var data;
        
        totalPayment = getTotalPayment(plainBody);
        
        while ((data = regExp.exec(plainBody)) !== null) { 
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
  
  if (rangeStatus != null) {
    rangeStatus.setValue("Concluído");
    rangeStatus.setBackground("#d9ead3");    
  }
  
  try {
    getNamedRangeFast(namedRanges, "StartDate").setValue((new Date()));
  } catch(err) {
    Logger.log(err);
  }
  
}

function getTotalPayment(messageBody){
  var regExp = new RegExp("Valor da transferência total: R\\$ ([\\d,\\.]*)\\.\\s+Dependendo de seu", "gi");
  var result = regExp.exec(messageBody);
  if (result == null) {
    return null;
  }
  return result[1];
}
