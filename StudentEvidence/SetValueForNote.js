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
 
 function setValueForNote(sheet, noteTag, newValue){
  Logger.log('setValueForNote(%s, %s, %s}', sheet, noteTag, newValue);
  var data = sheet.getDataRange();
  var notes = data.getNotes();
  
  var row = 0;
  var col = 0;
  
  var found = false;
  
  for( var r in notes){
    for ( var c in notes[r]){
      if( notes[r][c] == noteTag ){
        row = parseInt(r, 10) +1;
        col = parseInt(c, 10) +1;
        found = true;
        break;
      }
    }
    if( found) {
      break;
    }
  }
  
  var cell = data.getCell(row, col);
  cell.setValue(newValue);
}
