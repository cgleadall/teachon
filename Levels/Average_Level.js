/**
 * Get Average Level for a list - Must be a column list of levels.
 * @param {range} the list of levels to use to calculate the average.
 * @example AVERAGE_LEVEL(B3:B10)
 * @returns {String} Returns the average of the list of Levels
 * @customfunction

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
 
 function AVERAGE_LEVEL(list, show)
 {
  var result = "";
  var data = Array(list);
  
  var sum = 0;
  var valueCount = 0;
  var containsRorI = false;
  var extraInformation = "";
  
  var values = String(data).split(",");
  var numberOfValuesToAverage = values.length;
  
  var lookupArray = ["1-", "1", "1+", "2-", "2", "2+", "3-", "3", "3+", "4-", "4", "4+", "4++", "4+++"];

  for( var r = 0; r < numberOfValuesToAverage; r++)
  {
      var lookupValue = String(values[r]).toLowerCase();
      
      if( lookupValue.length > 0)
      {
        if( lookupValue == "r" || lookupValue == "i" || lookupValue == "4R")
        {      
          containsRorI = true;
          extraInformation += lookupValue.toUpperCase();
        }
        else
        {
          valueCount++
          sum += lookupArray.indexOf(lookupValue) + 1; // Need to add 1 because it is 0-indexed
        }
    }
  }

  var averageLevel = Math.round(sum/valueCount,0);

  if( averageLevel > 1 ){
    result = String(lookupArray[averageLevel-1]);  // Need to subtract 1 because it is 0-indexed
  }
  
  if(containsRorI && extraInformation.length < 5){  // add a '*' if there is an R or I in the range of levels
    result = result + "(" + extraInformation + ")";
  }
  else if ( containsRorI ) 
  {
    result = result + "*";
  }
  
  if(show){
    return result + " : " + JSON.stringify(values);
  } else {
    return result;
  }
}
 