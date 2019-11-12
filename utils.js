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
