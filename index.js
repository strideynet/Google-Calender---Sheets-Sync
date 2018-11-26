var CALENDAR_ID = ''
var DATE_RANGE = 2 * 365 * 24 * 60 * 60 * 1000 // 2 years

function syncDatabase () {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getActiveSheet()
  var calendar = CalendarApp.getCalendarById(CALENDAR_ID)

  var lastRow = sheet.getDataRange().getLastRow()
  var allData = sheet.getRange(2, 1, lastRow - 1, 4).getValues()

  var today = new Date()
  var calendarEvents = calendar.getEvents(new Date(today.getTime() - DATE_RANGE), new Date(today.getTime() + DATE_RANGE))
  var existingHolidays = {}

  for (var i = 0; i < calendarEvents.length; i++) {
    var event = calendarEvents[i]
    var id = event.getLocation()

    // Delete dupes
    if (existingHolidays[id]) {
      existingHolidays[id].deleteEvent()
    }

    existingHolidays[id] = event
  }

  for (var i = 0; i < allData.length; i++) {
    var employee = allData[i][0]
    if (employee) {
      var start = allData[i][1]
      var end = new Date(allData[i][2].setDate(allData[i][2].getDate() + 1))

      var title = employee + ' | ' + allData[i][3]

      // Delete events that already exist in order to update them
      if (existingHolidays[String(i + 2)]) {
        existingHolidays[String(i + 2)].deleteEvent()
      }

      calendar.createAllDayEvent(title, start, end, { description: 'reason: ' + allData[i][3], location: String(i + 2)}).removeAllReminders()
    }
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sync Calender').addItem('Start Sync', 'syncDatabase').addToUi()
}