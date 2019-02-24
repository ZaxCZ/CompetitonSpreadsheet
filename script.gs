/** @OnlyCurrentDoc */
function clearPlaneUsage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kalendarS = ss.getSheetByName("Kalendar");
  var letadlaS = ss.getSheetByName("Letadla");
  var nastaveniS = ss.getSheetByName("Nastaveni");

  var letadla  = letadlaS.getDataRange().getValues();
  var kalendarStart = nastaveniS.getDataRange().getValues()[1][0];  
  var lenghtOfCalendar = nastaveniS.getDataRange().getValues()[1][3];
  
  var column = 2; //B
  var row = 6;    //6
  var cell = kalendarS.getRange(row,column,letadla.length,lenghtOfCalendar+1);
  cell.setBackground("white");  
  cell.setValue('');
  cell.breakApart();
  cell.setBorder(null, true, null, true, true, null);
}

function drawPlaneUsage() {
  clearPlaneUsage();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kalendarS = ss.getSheetByName("Kalendar");
  var dataS = ss.getSheetByName("Data");
  var letadlaS = ss.getSheetByName("Letadla");
  var nastaveniS = ss.getSheetByName("Nastaveni");

  var letadla  = letadlaS.getDataRange().getValues();
  var data     = dataS.getDataRange().getValues();
  var kalendarStart = nastaveniS.getDataRange().getValues()[1][0];
  
  var kalendarStartTime = kalendarStart.getTime();
  
  for (var i = 0; i < letadla.length; i++) {
    var letadlo = letadla[i][0];
    for (var j = 1; j < data.length; j++) {
      if (letadlo == data[j][2]) {
        var conflictingData = dataS.getRange(j+1,1,1,6);
        conflictingData.setBackground("white"); // clean conflicts
        
        var start       = data[j][0];
        var end         = data[j][1];
        var name        = data[j][3]; 
        var surname     = data[j][4];
        var competition = data[j][5];
        var startTime   = start.getTime(), // time in ms
            endTime     = end.getTime();
        var duration      = Math.floor((endTime-startTime)/(24*3600*1000)) + 1; // in days
        var relativeStart = Math.floor((startTime-kalendarStartTime)/(24*3600*1000)); // in days
        
        var row    = 6 + i;             // start at 6
        var column = 2 + relativeStart; // start at B
        
        var cell = kalendarS.getRange(row,column,1,duration);
        
        if (cell.isPartOfMerge()) {
          // plane is taken
          conflictingData.setBackground("red"); // set conflict
        } else {
          cell.mergeAcross();
          cell.setBackground("#fcd2d2");
          cell.setValue(name + ' ' + surname + ' - ' + competition);
        }   
      }
    }
  }
}

// trigger
function onEdit(e) {
  // Set a comment on the edited cell to indicate when it was changed.
   
  var editS = e.range.getSheet();
  if (editS.getName() == "Data"){
    var range = e.range;
    range.setNote('Last modified: ' + new Date());
    drawPlaneUsage();
  }
  else return; 
}

//function onOpen(e){
//  drawPlaneUsage();
//}
