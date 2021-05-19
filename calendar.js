// Update the calendar for anything a month old up to 12 months from now
function getCalendarEvents(cal, startDate, endDate) {
  var options = {search: aosMainCalendarEventKey};
  
  var events = cal.getEvents(startDate, endDate, options);
  return events;
}

function getExistingEvents(existingCalEvents) {
  var output = [];
  for (var i = 0; i < existingCalEvents.length; i++) {
    var existing = existingCalEvents[i];
    var title = existing.getTitle();
    if (existing.isAllDayEvent()) {
      var start = existing.getAllDayStartDate();
      var end = existing.getAllDayEndDate();
      // adjust end date by 1 day so it matches with what is present in the spreadsheet
      end.setDate(end.getDate() - 1)
    } else {
      var start = existing.getStartTime();
      var end = existing.getEndTime();
    }
    var id = existing.getId();
    
    var obj = {title: title, start: start, end: end, id: id}
    output.push(obj);
  }
  return output;
}

function validExpectedEventDate(eventStart, eventEnd, syncStart, syncEnd) {
  if (eventStart > syncEnd) {
    return false;
  }
  if (eventEnd != "" && eventEnd < syncStart) {
    return false;
  }
  if (eventStart < syncStart && eventEnd == "") {
    return false;
  }
  return true;
}

function getExpectedEvents(sched, syncStart, syncEnd) {
  var output = [];
  var range = sched.getRange("A3:D");
  var numRows = range.getNumRows();
  var values = range.getValues();
  for (var i = 0; i < numRows; i++) {
    // event titles are in first column for "sprints"
    var title = values[i][0];
    if (title == "") {
      // and in second column for "important dates"
      title = values[i][1];
    }
    
    var start = values[i][2];
    var end = values[i][3];
    if (end == "") {
      end = new Date(start);
    }
    if (!validExpectedEventDate(start, end, syncStart, syncEnd)) {
      continue
    }
    // Accounted is a flag used to mark an expected event is already accounted for in the calendar
    output.push({title: title, start: start, end: end, accounted: false});
  }
  return output;
}

function sameEvent(existing, expected) {
  if (existing.title != expected.title) {
    return false;
  }
  if (expected.start.getTime() === existing.start.getTime() && expected.end.getTime() === existing.end.getTime()) {
    return true;
  }
  Logger.log("Different time for same event: %s: %s: %s; %s: %s", expected.title, expected.start, expected.end, existing.start, existing.end)
  return false;
}

function getEventsToDelete(expectedEvents, existingEvents) {
  var output = [];
  for (var i = 0; i < existingEvents.length; i++) {
    shouldDelete = true;
    var existing = existingEvents[i];
    for (var j = 0; j < expectedEvents.length; j++) {
      var expected = expectedEvents[j];
      if ((!expected.accounted) && sameEvent(existing, expected)) {
        shouldDelete = false;
        // The expected event is now accounted, any more dup entries must be deleted
        expected.accounted = true;
        break;
      }
    }
    if (shouldDelete) {
      Logger.log("Should Delete: %s: %s: %s: %s", existing.id, existing.title, existing.start, existing.end);
      output.push(existing)
    }
  }
  return output;
}

function deleteEvents(cal, eventsToDelete) {
  for (var i = 0; i < eventsToDelete.length; i++) {
    var toDelete = eventsToDelete[i];
    var event = cal.getEventById(toDelete.id);
    Logger.log("Deleting: %s: %s", toDelete.id, toDelete.title);
    event.deleteEvent();
  }
}

function getEventsToCreate(expectedEvents, existingEvents) {
  var output = [];
  for (var i = 0; i < expectedEvents.length; i++) {
    shouldCreate = true;
    var expected = expectedEvents[i];
    for (var j = 0; j < existingEvents.length; j++) {
      var existing = existingEvents[j];
      if (sameEvent(existing, expected)) {
        shouldCreate = false;
        break;
      }
    }
    if (shouldCreate) {
      Logger.log("Should create: %s: %s: %s", expected.title, expected.start, expected.end);
      output.push(expected);
    }
  }
  return output;
}

function createEvents(cal, eventsToCreate) {
  for (var i = 0; i < eventsToCreate.length; i++) {
    var toCreate = eventsToCreate[i];
    var options = {description: aosMainCalendarEventKey};
    Logger.log("Creating: %s: %s: %s", toCreate.title, toCreate.start, toCreate.end);
    if (toCreate.start.getTime() === toCreate.end.getTime()) {
      // If start and end time are the same, create an all day event.
//      cal.createEvent(toCreate.title, toCreate.start, toCreate.end, options)
      cal.createAllDayEvent(toCreate.title, toCreate.start, options)
    } else if ((toCreate.end.getDate()-toCreate.start.getDate())>0){
      // If it is an event spanning multiple days, create all-day event
      // end-date needs to be adjusted to the next day due to the way google calendar events work.
      toCreate.end.setDate(toCreate.end.getDate() + 1)

      cal.createAllDayEvent(toCreate.title, toCreate.start, toCreate.end, options)
    } else {
      // Single day event that has different start and end times
      cal.createEvent(toCreate.title, toCreate.start, toCreate.end, options); 
    }
    //cal.createEvent(toCreate.title, toCreate.start, toCreate.end, options); 
  }
}

function updateAOSMainCalendar(sched) {
  var cal = CalendarApp.getCalendarById(aosMainCalendarID);
  var startDate = new Date();
  // Start from a month ago from today
  startDate.setMonth(startDate.getMonth() - 1);
  var endDate = new Date();
  // End date is a year from today
  endDate.setYear(endDate.getYear() + 1);
  var existingCalEvents = getCalendarEvents(cal, startDate, endDate);
  var existingEvents = getExistingEvents(existingCalEvents);
  var expectedEvents = getExpectedEvents(sched, startDate, endDate);
  
  var eventsToDelete = getEventsToDelete(expectedEvents, existingEvents);
  deleteEvents(cal, eventsToDelete); 

  var toCreateEvents = getEventsToCreate(expectedEvents, existingEvents);
  createEvents(cal, toCreateEvents); 
}
