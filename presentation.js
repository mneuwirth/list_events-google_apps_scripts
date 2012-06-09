/*
 * A Google app script that adds your gcalendar events to a google spreadsheet.
 *
 * Copyright (c) 2011,2012 Emergya
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * Author: Francisco Perez <fperez@emergya.com>
 *
 */

// Presentation functions

function sync() {
  
  // Sheet by default
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Select Sheet1
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Sheet1"));
  var sheet = spreadsheet.getActiveSheet();

  // Get calendar user by Name 
  var cal = CalendarApp.getCalendarsByName("PUT_NAME_OF_CALENDAR")[0];
  //rowser.msgBox(Utilities.formatDate("06/05/2012", "dd/MM/yyyy"));
  // Get events in the date range
  var events = cal.getEvents(new Date("06/05/2012 09:00:00"),new Date("06/15/2012 09:00:00"));
 

  for (var i = 0; i <= events.length; ++i) {
    var row={};
    
    var event = events[i].getTitle();  // event
    var description = events[i].getDescription(); // description
    var list_emails  = ""; // list emails from
    
    var guests = events[i].getGuestsStatus();
    var contact;
    
    for (var j = 0; j < guests.length; ++j) {
      // Only getGuestStatus = YES or INVITED"
      // Find contact in emails address contacts
      contact = ContactsApp.findByEmailAddress(guests[j].getEmail());
      if (contact!=null &&  
          (guests[j].getGuestStatus()=="YES" || guests[j].getGuestStatus()=="INVITED")
         ){
          list_emails =   guests[j].getEmail() + " " + list_emails; 
      }
    }
     

    row[1]= event
    row[2]= description
    row[3]= list_emails; // emails from
    row[4]= events[i].getStartTime();  // start date
    row[5]= events[i].getEndTime();  // end date

 
    // Insert events data to spreedsheet
    for (var j = 1; j < 6; ++j) {
     
      sheet.getRange( i+1,j ).setValue(row[j]); 
      
     }

    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();

  }
  
    var app = UiApp.getActiveApplication();
    //app.close();

    //The following line is REQUIRED for the widget to actually close.
    return app;

}

