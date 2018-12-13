function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var result = getJson(sheet);


  var folder = DriveApp.getFolderById('XXXXXXXXXXXXXXXXXXXXXXXXXXX');
  createFile(folder, getSheetInfo(sheet).name, result)
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('スクリプト');
  menu.addItem('JSON変換', 'myFunction');
  menu.addToUi();
}

function createFile(folder, filename, obj) {
  var json = JSON.stringify(obj, null, 2);

  var fileName = filename+'.json';
  var contentType = 'application/json';
  var charset = 'utf-8';
  var fileItr = folder.getFilesByName(fileName);
  while (fileItr.hasNext()) {
    var file = fileItr.next();
    folder.removeFile(file);
  }
  var blob = Utilities.newBlob('', contentType, fileName).setDataFromString(json, charset);
  folder.createFile(blob);
}


function getJson(sheet, history) {
    if (typeof history === 'undefined')
        history = [];
    var info = getSheetInfo(sheet);

    if (history.indexOf(info.name) >= 0) {
        throw new Error('シートが循環参照になっています。')
    }
    history.push(info.name);

    if (info.type === 'Object') {
      return getObjectSheet(sheet, history)
    }
    if (info.type === 'Array') {
       return getArraySheet(sheet, history)
    }
    return  getObjectArraySheet(sheet, history)
}

function getObjectSheet(sheet, history) {
    var lastRow = sheet.getLastRow();
    var result = {};
    for (var i = 1; i <= lastRow; i++) {
        var colData = getColumnValues(sheet, i);
        result[colData[0]] = getValue(colData[1], history);
    }
    return result
}

function getArraySheet(sheet, history) {
    var lastRow = sheet.getLastRow();
    var result = [];
    for (var i = 1; i <= lastRow; i++) {
        var colData = getColumnValues(sheet, i);
        result.push(getValue(colData[0], history))
    }
    return result
}

function getObjectArraySheet(sheet, history) {
    var lastRow = sheet.getLastRow();
    var result = [];

    var keys = getColumnValues(sheet, 1);

    for (var i = 2; i <= lastRow; i++) {
        var obj = {};
        var values = getColumnValues(sheet, i);
        keys.forEach(function(key, i) {
            if (typeof values[i] !== 'undefined') {
                obj[key] = getValue(values[i], history);
            }
        });
        result.push(obj);
    }
    return result;
}

function getColumnValues(sheet, rowIndex) {
    var lastColumn = sheet.getLastColumn();
    return sheet.getRange(rowIndex, 1, rowIndex, lastColumn).getValues()[0]
}

function validateKeys(keys) {
    if (!keys.every(function(key) {
        return typeof key === 'string' && key.length > 0
    })) {
        Logger.log('不適切なキー名が存在します。')
    }
}

function isHashName(value) {
    return typeof value === 'string' && value.match(new RegExp("^{{#.+}}$"))
}

function getHashName(value) {
    if (typeof value !== 'string')
        return ''
    return value.replace(new RegExp("^{{#"), '').replace(new RegExp("}}$"), '');
}

function getValue(value, history) {
    var result = value;
    if (isHashName(result)) {
        var targetSheetName = getHashName(result);
        var pattern = new RegExp("^" + targetSheetName + "(\\?.*)*$");
        var sheets = SpreadsheetApp.getActive().getSheets();
        for (var i = 0; i < sheets.length; i++) {
            var sheetName = sheets[i].getName();
            if (sheetName.match(pattern) !== null) {
                result = getJson(sheets[i], history);
            }
        }
    }
    return result
}

function getSheetInfo(sheet) {
    var sheetName = sheet.getName();
    var name = sheetName;
    var type = '';
    var matched = sheetName.match(/\?.*/);
    if (matched !== null) {
        name = sheetName.substr(0, matched.index);
        type = sheetName.substr(matched.index + 1);
    }
    return {name: name, type: type}
}
