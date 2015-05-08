function onEdit() {
  var s = SpreadsheetApp.getActiveSheet();
  if( s.getName() == "Voting" ) { //checks that we're on the correct sheet
    var r = s.getActiveCell();
    if( r.getColumn() == 2 ) { //checks the column
      var nextCell = r.offset(0, 1);
      nextCell.setValue(new Date())
      var nextCell = r.offset(0, 2);
      nextCell.setValue(Session.getActiveUser().getEmail())
    }
  }
}
