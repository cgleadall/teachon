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
 
 function createStudentFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  
  ss.toast('Finding student list');
  
  var sheet = ss.getSheetByName('Student List');

  var ui = SpreadsheetApp.getUi();
  
  var parentFolderId = getParentFolderId(ssId);
  
  var data = sheet.getDataRange();
  var numRows = data.getNumRows();
  
  for( var r = 2; r <= numRows; r++){
    var studentCode = data.getCell(r, 4);
    var displayName = data.getCell(r, 5);
    
    if( studentCode.getValue().length == 0){
      Logger.log('studentCode are too short. Assume end of student list.');
      r = numRows;
      break;
    }
    var studentCodeValue = studentCode.getValue();
    var displayNameValue = displayName.getValue();
    ss.toast('Create sharable student file for \''+studentCodeValue+'\'');
    createStudentFile(ss, studentCodeValue, displayNameValue);    
    }

}


function getParentFolderId(ssId){

  var file = DriveApp.getFileById(ssId);
  var parents = file.getParents();
  if( parents.hasNext()){
    return parents.next().getId();
  }
  
  return null;
}


function createStudentFile(ss, studentCode, displayName){
  var sourceFileId = ss.getId();
  Logger.log('createStudentFile(%s, %s, %s)', sourceFileId, studentCode, displayName);

  var file = DriveApp.getFileById(sourceFileId);
  var folder = DriveApp.getFolderById(getParentFolderId(sourceFileId));
  
  var newStudentTemplateFilename = 'Student Template';
  var newFile = folder.getFilesByName(newStudentTemplateFilename).next();
  
  ss.toast('Creating sharable student template for: \'' + studentCode + '\'');
  var studentTemplateFile = newFile.makeCopy(studentCode, folder);
  
  var stSheet = SpreadsheetApp.openById(studentTemplateFile.getId());

  var sheet = stSheet.getActiveSheet();
  ss.toast('updating values on new sharable file');
  sheet.setName('Student\'s Assesment');
  setValueForNote(sheet, '#DisplayName', displayName);
  setValueForNote(sheet, '#StudentCode', studentCode);
  
}
