function convertStringToDate(str){
// str1 format should be dd/mm/yyyy
  var parts = str.split("/");
  var date = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  return date;
}

function addDays (date, days) {
  var dat = new Date(date.valueOf());
  dat.setDate(dat.getDate() + days);
  return dat;
}

function formatDate(data){
    var dia = data.getDate();
    if (dia.toString().length == 1)
      dia = "0"+dia;
    var mes = data.getMonth()+1;
    if (mes.toString().length == 1)
      mes = "0"+mes;
    var ano = data.getFullYear();  
    return dia+"/"+mes+"/"+ano;
}

function removeLineBreaksAndSpecialChars(emailText) {
  return emailText.replace(/\r?\n|\r/g," ").replace(/\*/g,"");
}

function getNamedRangeFast(namedRanges, name) {
  for(var i = 0; i<namedRanges.length; i++) {
    var namedRange = namedRanges[i];
    if (namedRange.getName() == name) {
      return namedRange.getRange();
    }
  }
}

function getStartDateFilterIfPresent(range) {
  try {
    var startDate = new Date(getNamedRangeFast(range, "StartDate").getValue());
    
    var startDateFilter = "after:" + startDate.toISOString().slice(0, 10);
    return startDateFilter;
  } catch (err) {}
  return "";
}


String.prototype.replaceAll = function(search, replacement) {
    var target = this;
    return target.replace(new RegExp(escapeRegExp(search), 'g'), replacement);
};

Date.prototype.addDays = function(days) {
  var dat = new Date(this.valueOf());
  dat.setDate(dat.getDate() + days);
  return dat;
}