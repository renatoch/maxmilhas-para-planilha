function onEditar(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var namedRanges = sheet.getNamedRanges();
  var rangeTrigger = getNamedRangeFast(namedRanges, "UpdateTrigger");
  
  if (rangeTrigger != null && rangeTrigger.getA1Notation() == e.range.getA1Notation()) {
    updateControlSpreadsheet();
  }
}
