function GetData() {
  clearSheet();
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2').activate();
  awalBaris = 2;
  tickerBaris = 2;
  while (spreadsheet.getRange('Ticker!$A'+tickerBaris).getValue() != "") {
    spreadsheet.getCurrentCell().setFormula('=query(GOOGLEFINANCE("IDX:"&indirect("Ticker!$A'+ tickerBaris +'",TRUE),"all",Ticker!$C$2,TODAY()),"select * label Col1\'\', Col2\'\', Col3\'\', Col4\'\', Col5\'\', Col6\'\'")');
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
    akhirBaris = spreadsheet.getCurrentCell().getRow();
    
    for (var i = awalBaris; i <= akhirBaris; i++) {
      spreadsheet.getRange('A'+i).setFormula('=Ticker!$A'+tickerBaris);
      /*spreadsheet.getRange('A'+i).setValue(spreadsheet.getRange('Ticker!$A'+tickerBaris).getValue());*/
    }
    
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    awalBaris = spreadsheet.getCurrentCell().getRow();
    tickerBaris = tickerBaris + 1;
  }
};

function clearSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};
