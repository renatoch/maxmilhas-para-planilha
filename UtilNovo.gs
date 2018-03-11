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