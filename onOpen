// ver 4.0   - add only Calendar = No and Status = Pending events.  auto-delete Yes & Done.  Skip Yes & Pending.
// future development - specify project in Type (sheet) into a var for calendarName.

function onOpen() {
var ui = SpreadsheetApp.getUi();
ui.createMenu('Sync to Calendar')
      .addItem('Add events', 'create_calendar')
      .addToUi();
}

function create_calendar() {
  
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var index = 2;  // 2nd row onwards down
var lastRow = sheet.getLastRow();

for (;index <= lastRow; index++){
  
  var type = sheet.getRange(index, 1, 1, 1).getValue();  //getRange(row, column, numRows, numColumns)
  var taskTitle = sheet.getRange(index, 2, 1, 1).getValue();
  var taskDesc = sheet.getRange(index, 3, 1, 1).getValue();
  var onCalendar = sheet.getRange(index, 4, 1, 1).getValue();
  var time = sheet.getRange(index, 5, 1, 1).getValue();
  var startDate = sheet.getRange(index, 6, 1, 1).getValue();
  var endDate = sheet.getRange(index, 7, 1, 1).getValue();
  var status = sheet.getRange(index, 8, 1, 1).getValue();
    
  if (onCalendar == 'Yes' && startDate && endDate && status == 'Done') // && means AND
  {
   
    var events =  CalendarApp.getCalendarsByName(type)[0].getEvents(startDate,endDate);
    delete_events(events);
    
  }
  
    else
      
    if (onCalendar == 'Yes' && startDate && endDate && status == 'Pending')
      
    {
      var events =  CalendarApp.getCalendarsByName(type)[0].getEvents(startDate,endDate);
      delete_events(events);
      
      var calendar = CalendarApp.getCalendarsByName(type)[0].createEvent(taskTitle,startDate,endDate,{description: taskDesc});
    
      
    }
  
  else
    
  var calendar = CalendarApp.getCalendarsByName(type)[0].createEvent(taskTitle,startDate,endDate,{description: taskDesc});
  sheet.getRange(index, 4, 1, 1).setValue('Yes');
  
  }
     
  
}
  

function delete_events(events) {
  
  for(var i=0; i<events.length;i++){
      var ev = events[i];
    
   // Browser.msgBox(ev.getTitle());
//      Logger.log(ev.getTitle()); // show event name in log
      ev.deleteEvent();
    }
  
}
