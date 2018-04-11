// Tailored for MITF
ENABLED = true;

DURCOL = "5";
INCOL = "3";
OUTCOL = "4";
DURCOLL = "E";
TOTAL = "TOTAL";
SCORECOLOR = "#ffffff";
TCRE = /\d\:\d\d\:\d\d\:\d\d/;
DURRE = /\d\:\d\d/;


function onEdit(e) {
  if (!ENABLED) return;

  switch (e.range.getColumn().toString()) {
    case DURCOL:
      updateTotalDuration();      
      break;
    case INCOL:
    case OUTCOL:
      updateCueLength(e);
      updateTotalDuration();
      break;
    default:
      // Do Nothing
  }
}





/////////////////////////////////////////////////////////
// Duration Sum by Color
function updateTotalDuration() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == TOTAL) {
        totalCell = sheet.getRange(i+1, j+2);
        totalCell.setValue(durationSumWithoutSource(DURCOLL+"1:"+DURCOLL+"100"));
      }
    }
  }
}

function durationSumWithoutSource(rangeSpecification) {
  var condition = function (cell) { return cell.getBackground() == SCORECOLOR };
  return sumByCondition(rangeSpecification, condition);
}

function sumByCondition(rangeSpecification, condition) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = sheet.getRange(rangeSpecification);
  
  var x = 0;
  
  for (var i = 1; i <= range.getNumRows(); i++) {
    for (var j = 1; j <= range.getNumColumns(); j++) {
      var cell = range.getCell(i, j);
      var cellVal = cell.getValue();

      if (condition(cell) && cellVal.match(DURRE)) {
        var parts = cellVal.split(":");
        x += (parseInt(parts[0], 10) * 60) + parseInt(parts[1], 10);
      }
    }
  }
  
  return durationStringFromSeconds(x);
}





/////////////////////////////////////////////////////////
// Diff Timecodes
function updateCueLength(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var row = e.range.getRow();
  var inCell = findConnectedCell(row, INCOL, -1);
  var outCell = findConnectedCell(row, OUTCOL, 1);
  var durationCell = sheet.getRange(outCell.getRow(), DURCOL);
  
  if (inCell === null || outCell === null) {
    // Give up.
    return false;
  }

  var diffText = diffTimecodes(inCell,outCell);
  
  if (diffText) {
    durationCell.setValue(diffText);
  }
}

function findConnectedCell(row, col, rowCheckInterval) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getRange(row,col);
  var val = cell.getValue();

  if (val.match(TCRE)) {
    return cell;
  }

  if (val.match(/connect/)) {
    return findConnectedCell(row + rowCheckInterval, col, rowCheckInterval);
  }
  
  return null;
}

function diffTimecodes(inCell, outCell) {
  var inVal = inCell.getValue();
  var outVal = outCell.getValue();

  var x = 0;

  if (inVal.match(TCRE) && outVal.match(TCRE)) {
    // Ignoring frame-level precision
    var inSeconds = getSecondsFromTimecode(inVal);
    var outSeconds = getSecondsFromTimecode(outVal);

    var diff = outSeconds - inSeconds;
    return durationStringFromSeconds(diff);
  }

  return false;
}

function getSecondsFromTimecode(tc) {
  var parts = tc.split(":");
  return (parseInt(parts[0], 10) * 60 * 60) + (parseInt(parts[1], 10) * 60) + (parseInt(parts[2], 10));
}





/////////////////////////////////////////////////////////
// Utility
function durationStringFromSeconds(seconds) {
  var min = Math.floor(seconds/60);
  var sec = seconds % 60;

  var minStr = min < 10 ? "0" + min : min;
  var secStr = sec < 10 ? "0" + sec : sec;

  return "" + minStr + ":" + secStr;
}





/////////////////////////////////////////////////////////
// Debug
function debugIntoCell(row, col, text) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.getActiveSheet().getRange(row, col).setValue(text);
}
