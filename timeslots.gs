function findCommonFreeTime() {
  // Configuration
  const CALENDARS = [
    { 
	// Calendar IDs obscured
    }
  ];
  
  const DAYS_TO_CHECK = 7;
  const SLOT_DURATION = 30;
  
  const WEEKDAY_RANGES = [
    { start: 07, end: 23 }
  ];
  
  const WEEKEND_RANGES = [
    { start: 07, end: 23 }
  ];
  
  const COLORS = {
    BUSY: '#FF9999',
    FREE: '#99FF99',
    COMMON_FREE: '#00FF00',
    HEADER: '#E6E6E6'
  };
  
  const ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1JejWq0H90CR5CJ26A7Kpf2UsqUyqTQV_bIPP7wkVOzA/edit?gid=2092247677#gid=2092247677');
  const sheet = ss.getSheetByName('Availability Grid') || ss.insertSheet('Availability Grid');
  
  try {
    const startDate = new Date();
    startDate.setHours(0, 0, 0, 0);
    const endDate = new Date(startDate.getTime() + (DAYS_TO_CHECK * 24 * 60 * 60 * 1000));
    
    const calendarObjects = [];
    const errorMessages = [];
    
    for (const calendarConfig of CALENDARS) {
      try {
        let calendar;
        if (calendarConfig.id === 'primary') {
          calendar = CalendarApp.getDefaultCalendar();
        } else {
          calendar = CalendarApp.getCalendarById(calendarConfig.id);
        }
        
        if (!calendar) {
          errorMessages.push(`Could not access calendar for ${calendarConfig.name}`);
          continue;
        }
        
        calendarObjects.push({
          calendar: calendar,
          name: calendarConfig.name
        });
        
      } catch (e) {
        errorMessages.push(`Error accessing calendar for ${calendarConfig.name}: ${e.toString()}`);
      }
    }
    
    if (calendarObjects.length === 0) {
      throw new Error('No calendars could be accessed. ' + errorMessages.join('; '));
    }
    
    sheet.clear();
    
    const headers = ['Date', 'Time'];
    calendarObjects.forEach(({name}) => headers.push(name));
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(COLORS.HEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    const calendarEvents = [];
    for (const {calendar} of calendarObjects) {
      try {
        const events = calendar.getEvents(startDate, endDate);
        calendarEvents.push(events);
      } catch (e) {
        Logger.log(`Error getting events: ${e.toString()}`);
        calendarEvents.push([]);
      }
    }
    
    let currentRow = 2;
    for (let d = 0; d < DAYS_TO_CHECK; d++) {
      const currentDate = new Date(startDate.getTime() + (d * 24 * 60 * 60 * 1000));
      const dayOfWeek = currentDate.getDay();
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      const timeRanges = isWeekend ? WEEKEND_RANGES : WEEKDAY_RANGES;
      
      for (const range of timeRanges) {
        for (let hour = range.start; hour < range.end; hour++) {
          for (let minute = 0; minute < 60; minute += SLOT_DURATION) {
            const slotStart = new Date(currentDate.getTime());
            slotStart.setHours(hour, minute, 0, 0);
            
            const slotEnd = new Date(slotStart.getTime() + SLOT_DURATION * 60000);
            
            const rowData = [
              Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
              Utilities.formatDate(slotStart, Session.getScriptTimeZone(), 'HH:mm')
            ];
            
            const availability = [];
            calendarEvents.forEach(events => {
              const isBusy = events.some(event => {
                const eventStart = event.getStartTime();
                const eventEnd = event.getEndTime();
                return slotStart < eventEnd && slotEnd > eventStart;
              });
              availability.push(isBusy);
            });
            
            availability.forEach(() => rowData.push(''));
            
            const range = sheet.getRange(currentRow, 1, 1, rowData.length);
            range.setValues([rowData]);
            
            for (let i = 0; i < availability.length; i++) {
              const cell = sheet.getRange(currentRow, i + 3);
              cell.setBackground(availability[i] ? COLORS.BUSY : COLORS.FREE);
            }
            
            if (availability.every(isBusy => !isBusy)) {
              range.setBackground(COLORS.COMMON_FREE);
            }
            
            currentRow++;
          }
        }
      }
    }
    
    sheet.autoResizeColumns(1, headers.length);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    
    const legendRow = currentRow + 1;
    sheet.getRange(legendRow, 1).setValue('Legend:');
    sheet.getRange(legendRow, 2).setValue('Busy').setBackground(COLORS.BUSY);
    sheet.getRange(legendRow, 3).setValue('Free').setBackground(COLORS.FREE);
    sheet.getRange(legendRow, 4).setValue('Common Free Time').setBackground(COLORS.COMMON_FREE);
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    sheet.getRange('A1').setValue('Error: ' + error.toString());
  }
}

function createTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    
    ScriptApp.newTrigger('findCommonFreeTime')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
      
    SpreadsheetApp.getActiveSpreadsheet().toast('Daily update trigger created successfully', 'Success');
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error creating trigger: ' + error.toString(), 'Error');
  }
}

