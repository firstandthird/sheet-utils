function sheetToConfig(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = sheet.getSheetByName(sheetName);

  const configRange = configSheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());

  const sheetValues = configRange.getValues();
  const config = [];
  // For Horiz - assume 0 is headers, and 1 is values.
  var hdrRange = sheetValues.shift();
  sheetValues.forEach(function(valRange) {
    var _config = {};
    for (var i = 0; i < hdrRange.length; i++) {
      var key = hdrRange[i].toLowerCase().replace(' ', '_');
      _config[key] = valRange[i];
    }
    config.push(_config);
  });
  return config;
}

// Old style ... vertical configs for readability.
function sheetToConfigVertical(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = sheet.getSheetByName(sheetName);
  const configRange = configSheet.getRange(1, 1, configSheet.getLastRow(), configSheet.getLastColumn());

  const sheetValues = configRange.getValues();
  const config = [];
  // For vertical - key is 0 in each row
  sheetValues.forEach(function(r) {
    var key = r.shift().toLowerCase().replace(' ', '_');
    r.forEach(function(v, k) {
      if (!config[k]) {
        config[k] = {};
      }
      config[k][key] = v;
    });
  });

  return config;
}

function cellsToTable(cells) {
  var valArr = [];
  valArr.push('<table><thead><tr>');
  const hdr = cells.shift();
  const hdrVals = [];
  hdr.forEach(function(c) {
    hdrVals.push('<th>' + c + '</th>');
  });

  valArr.push(hdrVals.join(''));
  valArr.push('</tr></thead><tbody>');

  cells.forEach(function(r) {
    valArr.push('<tr>');
    r.forEach(function(c) {
      valArr.push('<td>' + c + '</td>');
    });
    valArr.push('</tr>');
  });
  valArr.push('</tbody></table>');

  return valArr;
}

function showScheduler() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  
  if (triggers.length) {
    Logger.log(triggers[0].getUniqueId()); 
  }
  
  var template = HtmlService.createTemplateFromFile('Index');
  template.hasTriggers = (triggers.length > 0);
  if (template.hasTriggers) {
    var trigger = triggers[0];
    var source = trigger.getTriggerSource();
  }
  var html = template.evaluate()
    .setWidth(700);
  
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Dialog title');
}


function updateSchedule(data) {
  var upData = {};
  data.forEach(function(v){
    upData[v.name] = v.value;
  });
  // Delete the triggers by default.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var weekDay = null;
  switch(upData.dayOfWeek) {
    case '0':
      weekDay = ScriptApp.WeekDay.MONDAY;
      break;
    case '1':
      weekDay = ScriptApp.WeekDay.TUESDAY;
      break;
    case '2':
      weekDay = ScriptApp.WeekDay.WEDNESDAY;
      break;
    case '3':
      weekDay = ScriptApp.WeekDay.THURSDAY;
      break;
    case '4':
      weekDay = ScriptApp.WeekDay.FRIDAY;
      break;
    case '5':
      weekDay = ScriptApp.WeekDay.SATURDAY;
      break;
    case '6':
      weekDay = ScriptApp.WeekDay.SUNDAY;
      break;
  }
  // Set up a new trigger?
  if (upData.automate) {
    switch (upData.interval) {
      case 0:
      case '0':
        // Every Hour
        ScriptApp.newTrigger('fetchAll')
          .timeBased()
          .everyHours(1)
          .create();
        break;
      case 1:
      case '1':
        // Every Day
        ScriptApp.newTrigger('fetchAll')
          .timeBased()
          .everyDays(1)
          .atHour(upData.hourOfDay)
          .create();
        break;
      case 2:
      case '2':
        // Every Week
        ScriptApp.newTrigger('fetchAll')
          .timeBased()
          .everyWeeks(1)
          .onWeekDay(weekDay)
          .atHour(upData.hourOfDay)
          .create();
        break;
      case 3:
      case '3':
        // Every Month
        ScriptApp.newTrigger('fetchAll')
          .timeBased()
          .onMonthDay(upData.dayOfMonth)
          .atHour(upData.hourOfDay)
          .create();
        break;
      default:
        Logger.log('Invalid interval');
    }
  }
}
