var ETICKETCOLUMN = 7;
var ESTIMATEDRECEIVEDATECOLUMN = 5;
var SALEVALUEPERMILECOLUMN = 3;
var SALEVALUE = 4;

function updateControlSpreadsheet() {
  var sheet, gmailThreads,firstMessageSubject, firstMessage, firstMessagePlainBody, saleWasCancelled, sucessfulSale, i,
      method, date, transactionCode, eTicket, account, airline, airmilesAmount, saleValue, saleValuePerMile, estimatedReceiveDate, boardingFee, luggageFee;

  method = "MaxMilhas"
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Import");
  if (sheet == null) {
    throw new Error("There must be an 'Import' tab on the spreadsheet");
  }
  var namedRanges = sheet.getNamedRanges();
  var rangeStatus = getNamedRangeFast(namedRanges, "Status");

  var startDateFilter = getStartDateFilterIfPresent(namedRanges);
  
  if (rangeStatus != null) {
    rangeStatus.setValue("Getting Gmail threads...");
    rangeStatus.setBackground("#f9cb9c");
  }
  
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
    firstMessageSubject = gmailThreads[i].getFirstMessageSubject();

    if (firstMessageSubject.search("Venda de milhas - código: ")>-1 && firstMessageSubject.search("Re:")==-1) {
      firstMessage = gmailThreads[i].getMessages()[0];
      firstMessagePlainBody = removeLineBreaksAndSpecialChars(firstMessage.getPlainBody());
      
      saleWasCancelled = firstMessagePlainBody.indexOf("cancelamento") > -1;
      
      if(!saleWasCancelled){
        
        var splittedSubject = firstMessageSubject.split(" - ");
        
        eTicket = splittedSubject[2].replace("e-ticket: ","").trim();
          
        if( !(eTicket in eticketsMapping)){
          date = getDate(firstMessagePlainBody);
          transactionCode = splittedSubject[1].replace("código: ","").trim()
          account = getAccount(firstMessagePlainBody);
          airline = getAirline(firstMessagePlainBody);
          airmilesAmount = getAirmilesAmount(firstMessagePlainBody);
          saleValue = getSaleValue(firstMessagePlainBody);
          saleValuePerMile = getSaleValuePerMile(firstMessagePlainBody);
          boardingFee = getBoardingFee(firstMessagePlainBody);
          luggageFee = getLuggageFee(firstMessagePlainBody);
          estimatedReceiveDate = Utilities.formatDate(addDays(convertStringToDate(date), 20), "America/Sao_Paulo", "dd/MM");
          
          sheet.appendRow([date, airline, airmilesAmount, saleValuePerMile, saleValue, estimatedReceiveDate, transactionCode, eTicket, account, method, boardingFee, luggageFee]);
          sheet.getRange(sheet.getLastRow(), SALEVALUEPERMILECOLUMN + 1).setNumberFormat("R$ #,##0.00;R$ (#,##0.00)");
          sheet.getRange(sheet.getLastRow(), SALEVALUE + 1).setNumberFormat("R$ #,##0.00;R$ (#,##0.00)");
          eticketsMapping[eTicket] = true;
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

function getEticketsMapping(sheet) {
  var map = {};
  var data = sheet.getDataRange().getValues(); // read all data in the sheet
  
  for(n=0;n<data.length;++n){
    map[data[n][ETICKETCOLUMN].toString()] = n + 1;
  }
  return map;
}
        
function getSubstringInTheMiddle(originalString, first, second){
  var lastIndexOfFirstString = originalString.lastIndexOf(first)+first.length;
  var firstIndexOfSecondString = originalString.lastIndexOf(second);
  var substringInTheMiddle = "";
  if(lastIndexOfFirstString != -1 && firstIndexOfSecondString != -1){
    substringInTheMiddle = originalString.substring(lastIndexOfFirstString,firstIndexOfSecondString).trim();
  }
  return substringInTheMiddle;
}

function getDate(messageBody){
  var previousString = "vendidas no dia ";
  var firstIndexOfDesiredString = messageBody.lastIndexOf(previousString)+previousString.length; 
  var lastIndexOfDesiredString = firstIndexOfDesiredString + 10;
  return messageBody.substring(firstIndexOfDesiredString,lastIndexOfDesiredString).trim()
}

function getAccount(messageBody){
  
  var fullName = "";

  var firstIndexOfSecondString = messageBody.lastIndexOf("Ficamos felizes em");
  
  if(firstIndexOfSecondString != -1){
    fullName = messageBody.substring(0,firstIndexOfSecondString).trim();
    
    var firstString = "Olá, ";
    
    var firstStringFirstIndex = fullName.lastIndexOf(firstString);
    
    if(firstStringFirstIndex == -1){
      firstString = "Caro(a), "
      firstStringFirstIndex = fullName.lastIndexOf(firstString);
    }
    
    var lastIndexOfFirstString = firstStringFirstIndex + firstString.length;
    fullName = fullName.substring(lastIndexOfFirstString,fullName.length)
    fullName = fullName.trim()
  }
  
  var firstName = fullName.split(" ")[0];
  return firstName;
}

function getAirline(messageBody){
    var airline = getSubstringInTheMiddle(messageBody, "da sua oferta ", " foram");
    var membershipProgramAndAirline = "Erro";
    if(airline == "Avianca"){
      membershipProgramAndAirline = "Avianca";
    }else if(airline == "Azul"){
      membershipProgramAndAirline = "Azul";
    }else if(airline == "Latam"){
      membershipProgramAndAirline = "Multiplus";
    }else if(airline == "Gol"){
      membershipProgramAndAirline = "Smiles";
    }
    
  return membershipProgramAndAirline;
}

function getAirmilesAmount(messageBody){
  
  var airmilesAmount = "Erro"
  
  if(messageBody.indexOf("comunicar que ") >-1){
    airmilesAmount = getSubstringInTheMiddle(messageBody, "comunicar que ", " milhas da sua oferta")
  }else if(messageBody.indexOf("comunicá-lo que ") >-1){
    airmilesAmount = getSubstringInTheMiddle(messageBody, "comunicá-lo que ", " milhas* da sua oferta") 
  }
  
  return airmilesAmount;
}

function removeAllButNumbers(originalString) {
  return originalString.replace(/[^0-9.]/g, "");
}

function getSaleValue(messageBody){
  return getSubstringInTheMiddle(messageBody, "o valor de ", " referente as")
}


function getSaleValuePerMile(messageBody){
  return getSubstringInTheMiddle(messageBody, " (R$", " cada 1.000")
}

function getBoardingFee(messageBody){
  var boardingFee = "";

  var firstIndexOfSecondString = messageBody.lastIndexOf(" pela taxa de embarque");
  
  if(firstIndexOfSecondString != -1){
    boardingFee = messageBody.substring(0,firstIndexOfSecondString).trim();
    
    var lastIndexOfFirstString = boardingFee.lastIndexOf("R$ ")+ "R$ ".length;
    boardingFee = boardingFee.substring(lastIndexOfFirstString,boardingFee.length)
  }
  
  return boardingFee.trim();
}

function getLuggageFee(messageBody){
  return getSubstringInTheMiddle(messageBody, "embarque e "," pela bagagem ")
}
