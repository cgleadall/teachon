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
 
 function createNewCourse() {
  var ui = SpreadsheetApp.getUi();
  var masterEvidenceFilename = 'Student Evidence Master';
    
  var response = ui.prompt('Create a New Class/Course', 'Please enter the \'Name\' or \'Course Code\' for this new course', ui.ButtonSet.OK_CANCEL);
  if( response.getSelectedButton() == ui.Button.OK )
  {
    var newCourseName = response.getResponseText();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssId = ss.getId();
    ss.toast('Collecting information needed to create a new class/course');
    
    var parentFolderId = getParentFolderId(ssId);
    
    var file = DriveApp.getFileById(ssId);
    var folder = DriveApp.getFolderById(parentFolderId);

    var newFolder = folder.createFolder(newCourseName);
    Logger.log(newFolder);
    
    var contents = folder.getFilesByName(masterEvidenceFilename);
    var file;
    var name;

    var message = '';
    
    while(contents.hasNext()) {
      file = contents.next();
      name = file.getName();
      Logger.log('filename: ' +name);

      if (name == masterEvidenceFilename ){
        var newFilename = name.replace(/Master/, newCourseName);
        ss.toast('Creating new Master sheet for course: \'' + newFilename + '\'');
        file.makeCopy(newFilename, newFolder);
        
        var newStudentTemplateFilename = 'Student Template';
        ss.toast('Creating new Student Template file for course: \'' + newStudentTemplateFilename + '\'');
        var studentTemplateFile = file.makeCopy(newStudentTemplateFilename, newFolder);
        var stSheet = SpreadsheetApp.openById(studentTemplateFile.getId());
        stSheet.deleteSheet(stSheet.getSheetByName('Student List'));
        stSheet.setActiveSheet(stSheet.getSheetByName('Master'));
        stSheet.getActiveSheet().setName('Student\'s Evaluation');
        
      }
    }
  }
}

