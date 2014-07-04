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

 /**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
 
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Student Evidence')
      .addSubMenu(ui.createMenu('Administration')
        .addItem('Create New Class/Course', 'createNewCourse'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Set-up')
          .addItem('Create Student Sheets', 'createStudentSheets')
          .addItem('Create Student Files', 'createStudentFiles'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Synchronize')
        .addItem('Send Current Sheet to Student\'s share', 'synchornizeCurrentSheet')
        .addItem('Update All sheets to Students\' shares', 'synchronizeAllStudentSheets'))
      //.addItem('Files in Folder', 'listFiles')
      .addToUi();

};

function listFiles(id){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var ui = SpreadsheetApp.getUi();
  ui.alert(ssId);
  
  //var folder = DriveApp.getFolderById(ssId);
  var file = DriveApp.getFileById(ssId);
  var folder = DriveApp.getFolderById(file.getParents().next().getId())
  
  var contents = folder.getFiles();
  var file;
  var name;

  var message = '';
  
  while(contents.hasNext()) {
    file = contents.next();
    name = file.getName();
    message = message + ', ' + name;
  }
  
  ui.alert(message);
  
  
};

