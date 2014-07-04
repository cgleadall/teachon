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
 
 function synchornizeCurrentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getFolderById(getParentFolderId(ss.getId()));
  
  synchronizeStudentSheet(ss.getActiveSheet().getName(), folder);
}


function synchronizeAllStudentSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Student List');
  
  var data = sheet.getDataRange();
  var numRows = data.getNumRows();

  var folder = DriveApp.getFolderById(getParentFolderId(ss.getId()));


  for( var r = 2; r <= numRows; r++){
    var studentCode = data.getCell(r, 4);
    
    if( studentCode.getValue().length == 0){
      Logger.log('studentCode is too short. Assume end of student list.');
      r = numRows;
      break;
    }
    var studentCodeValue = studentCode.getValue();

    synchronizeStudentSheet(studentCodeValue, folder);
  }

}


function synchronizeStudentSheet(studentCode, folder){
  Logger.log('synchronizeStudentSheet(%s, %s)', studentCode, folder);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.toast('Sending your assesment to shared file for: \''+studentCode+'\'');
  
  var studentFiles = folder.getFilesByName(studentCode);
  if( studentFiles.hasNext() )
  {
    var file = studentFiles.next();
    Logger.log('About to open spreadsheet file by Id ' + file.getId());
    var targetSpreadsheet = SpreadsheetApp.openById(file.getId());

    var sourceSheet = ss.getSheetByName(studentCode);
    Logger.log('Source sheet : ' + sourceSheet.getName());
    
    var newTeacherSheet = sourceSheet.copyTo(targetSpreadsheet);
    Logger.log('Create new teacher sheet');
    
    Logger.log(targetSpreadsheet);
    var previousTeacherSheet = targetSpreadsheet.getSheetByName('Teacher\'s Assesment');
    
    if( previousTeacherSheet != null ){
      targetSpreadsheet.deleteSheet(previousTeacherSheet);
      Logger.log('deleted previous teacher sheet');
    }
    newTeacherSheet.setName('Teacher\'s Assesment');
    Logger.log('renamed new teach sheet');
  }
}
