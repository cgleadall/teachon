/**
    Copyright (C) 2014  Craig Gleadall
    https://github.com/cgleadall/teachon

    This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License along
    with this program; if not, write to the Free Software Foundation, Inc.,
    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
 */
 
 function createStudentSheets() {
  var spreadsheet = SpreadsheetApp;

  var ui = spreadsheet.getUi();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Student List');
  
  var data = sheet.getDataRange();
  var numRows = data.getNumRows();
  
  for( var r = 2; r <= numRows; r++){
    var studentCode = data.getCell(r, 4);
    var displayName = data.getCell(r, 5);
    
    if( studentCode.getValue().length == 0 || displayName.getValue().length == 0){
      Logger.log('studentCode and displayName are too short. Assume end of student list.');
      r = numRows;
      break;
    }
    var studentCodeValue = studentCode.getValue();
    ss.toast('Creating new sheet for ' + studentCodeValue);
    createStudentSheet(displayName.getValue(), studentCodeValue, 'Master');    
  }
};


function createStudentSheet(displayName, studentCode, sourceSheetName){
  Logger.log('createStudentSheet(%s, %s, %s}', displayName, studentCode, sourceSheetName);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);

  var newSheet = sourceSheet.copyTo(ss);
  newSheet.setName(studentCode);
  
  setValueForNote(newSheet, '#DisplayName', displayName);
  setValueForNote(newSheet, '#StudentCode', studentCode);
  
}
