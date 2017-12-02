var DIALOG_TITLE = 'Kuukausiohjelma';

// Spreadsheet id for containing Field Ministry Meetings
var SRC_MEETINGS = 'spreadsheet-id';
var SRC_SUNDAY_MEETINGS = 'spreadsheet-id';
var SRC_WEEKLY_TASKS = 'spreadsheet-id';
var SRC_FIELD_MINISTRY = 'spreadsheet-id';
var TARGET_FOLDER = 'folder-id';

var DATE_COLUMN = 0;

// User will select these with UI
var YEAR;
var MONTH;

// matches with "{{ for-row foo.bar[0] as row | filter:argument }}"
var EXPRESSION_REG_EX = '{{[^}]*}}';
var exprSearchResult = null;


/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Tee kuukausiohjelma', 'showDialog')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 */
function showDialog() {
  var ui = HtmlService.createTemplateFromFile('Dialog')
      .evaluate()
      .setWidth(400)
      .setHeight(200);
  DocumentApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}

function debugEntryPoint() {
  createScheduleDoc('2016', '1');
}

/**
 * Entry point. Creates a new document by rendering template document with
 * input data.
 *
 * @param {String} year Year for the schedule in 4 digits.
 * @param {String} month Month for the schedule (1-12).
 */
function createScheduleDoc(year, month, template) {
  YEAR = year;
  MONTH = month;
  var input = readInput();
  var newDocName = 'Kuukausiohjelma ' + year + '-';
  if (month < 10) {
    newDocName += '0';
  }
  newDocName += month;
  var doc = copyDocument(template, newDocName);
  render(doc, input);
  return newDocName;
}

/************************************************
 * JSON BUILDER
 ************************************************/

/**
 * Reads input data from defined input documents.
 *
 * @return {Object} The object containing input values.
 */
function readInput() {
  var input = {};
  var mondays = findSplitDaysInMonth(parseInt(YEAR), parseInt(MONTH));
  Logger.log("HEI");
  Logger.log(mondays);
  var monthFilter = function(date) {
    // dates between first monday and last monday
    return (date >= mondays[0] && date < mondays[mondays.length - 1]);
  };
  input['sun'] = readSpreadsheetInput(SRC_SUNDAY_MEETINGS, monthFilter);
  input['field'] = readSpreadsheetInput(SRC_FIELD_MINISTRY, monthFilter);
  input['tasks'] = readSpreadsheetInput(SRC_WEEKLY_TASKS, monthFilter);
  input['meetings'] = readSpreadsheetInput(SRC_MEETINGS, monthFilter);
  //input['field'] = readFieldMinistryInput(monthFilter);
  input['weeks'] = getWeeklyInput(input, mondays);
  input = sortWeeklyFields(input);
  input['month'] = monthAsText(MONTH);
  input['year'] = YEAR;
  Logger.log('WEEKS');
  Logger.log(input['weeks']);
  return input;
}

//function readFieldMinistryInput(monthFilter) {
//  var input = {};
//  var ss = SpreadsheetApp.openById(SRC_FIELD_MINISTRY);
//  var sheets = ss.getSheets();
//  for (var i = 0; i < sheets.length; i++) {
//    input[sheets[i].getName()] = filterSheetRows(sheets[i], monthFilter);
//  }
//  return input;
//}

//function readMeetingsInput(monthFilter) {
//  var input = {};
//  var ss = SpreadsheetApp.openById(SRC_MEETINGS);
//  var sheets = ss.getSheets();
//  for (var i = 0; i < sheets.length; i++) {
//    input[sheets[i].getName()] = getMonthRows(sheets[i], monthFilter);
//  }
//  return input;
//}

function readSpreadsheetInput(spreadsheetId, monthFilter) {
  var input = {};
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    input[sheets[i].getName()] = filterSheetRows(sheets[i], monthFilter);
  }
  Logger.log(input);
  return input;
}

function testi3() {
  YEAR = '2016';
  MONTH = '7';
  var input = {};
  var mondays = findSplitDaysInMonth(parseInt(YEAR), parseInt(MONTH));
  var monthFilter = function(date) {
    // dates between first monday and last monday
    return (date >= mondays[0] && date < mondays[mondays.length - 1]);
  };
  Logger.log(mondays);
  //input['field'] = readFieldMinistryInput(monthFilter);
  //input['weeks'] = getWeeklyInput(input, mondays);
  //input = sortWeeklyFields(input);
}

function sortWeeklyFields(input) {
  for (var i = 0; i < input['weeks'].length; i++) {
    var fieldObj = input['weeks'][i]['field'];
    var weeklyAll = [];
    for (var group in fieldObj) {
      for (var j = 0; j < fieldObj[group].length; j++) {
        weeklyAll.push({ group: group, meeting: fieldObj[group][j] });
      }
    }
    weeklyAll.sort(function(a, b) {
      if (a.meeting[DATE_COLUMN] < b.meeting[DATE_COLUMN]) {
        return -1;
      }
      else if (a.meeting[DATE_COLUMN] < b.meeting[DATE_COLUMN]) {
        return 1;
      }
      return 0;
    });
    input['weeks'][i]['fieldSorted'] = weeklyAll;
  }
  return input;  
}

function getWeeklyInput(input, mondays) {
  var weekRows = [];
  for (var i = 0; i < mondays.length - 1; i++) {
    //Logger.log(i);
    var allEmpty = true;
    var obj = {};
    for (var category in input) {
      obj[category] = {};
      for (var prop in input[category]) {
        //Logger.log(prop);
        //Logger.log(meetings[prop]);
        obj[category][prop] = [];
        for (var j = 0; j < input[category][prop].length; j++) {
          var date = input[category][prop][j][DATE_COLUMN];
          if (date >= mondays[i] && date < mondays[i + 1]) {
            obj[category][prop].push(input[category][prop][j]);
            allEmpty = false;
          }
        }
      }
    }
    if (!allEmpty) {
      obj.start = mondays[i];
      var endDate = new Date(mondays[i + 1].getFullYear(), mondays[i + 1].getMonth(), mondays[i + 1].getDate() - 1, 0, 0, 0);
      obj.end = endDate;
      weekRows.push(obj);
    }
  }
  return weekRows;
}

function testi() {
  var input = {};
  YEAR = '2015';
  MONTH = '11';
  var ss = SpreadsheetApp.openById(SRC_MEETINGS);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    //var foo = getMonthRows(sheets[i]);
    //var foo = filterSheetRows(sheets[i], function(date) {
    //  var isMonth = isSelectedMonth(date);
    //});
    var foo = getMonthRowsByWeek(sheets[i]);
    Logger.log(foo);
  }
}

/************************************************
 * DOCUMENT BUILDER
 ************************************************/

/**
 * Render document with given input.
 *
 * @param {Object} input Input values for rendering.
 * @param {Document} doc Document to be rendered.
 */
function render(doc, input) {
  var range = getFullRange(doc);
  var expressions = findExpressions(range, true);
  Logger.log(expressions);
  renderExpressions(doc, expressions, input);
}

function ItemTree(parent) {
  this.parent = parent;
  this.items = [];

  this.add = function(item) {
    this.items.push(item);
  };

  this.createChild = function() {
    var child = new ItemTree(this);
    return child;
  };

  this.endChild = function() {
    this.parent.items.push(this.items);
    return this.parent;
  };

  this.asArray = function() {
    return this.items;
  };

  this.debug = function() {
    Logger.log(this.items);
  };
}

/**
 * Returns Range of full document containing all Text elements.
 *
 * @return {Range}
 */
function getFullRange(doc) {
  var body = doc.getBody();
  var textElements = getTextElements(body);
  var rangeBuilder = doc.newRange();
  for (var i = 0; i < textElements.length; i++) {
    rangeBuilder.addElement(textElements[i]);
  }
  return rangeBuilder.build();
}

function getTextElements(body) {
  var elements = [];
  var searchResult = null;
  while (searchResult = body.findElement(DocumentApp.ElementType.TEXT, searchResult)) {
    elements.push(searchResult.getElement());
  }
  return elements;
}

function isInsideRangeElement(innerRangeElement, outerRangeElement) {
  if (innerRangeElement.isPartial() && outerRangeElement.isPartial()) {
    if ((innerRangeElement.getStartOffset() < outerRangeElement.getStartOffset()) ||
      (innerRangeElement.getEndOffsetInclusive() > outerRangeElement.getEndOffsetInclusive())) {
      return false;
    }
  }
  else if (!innerRangeElement.isPartial() && outerRangeElement.isPartial()) {
    return false;
  }
  return true;
}

/**
 * Finds all {{ expression }} from given range.
 *
 * @return {Array} Array of expression objects.
 */
function findExpressions(range) {
  var parent = null;
  var expressionTree = new ItemTree();
  var rangeElements = range.getRangeElements();
  for (var i = 0; i < rangeElements.length; i++) {
    var rangeElement = rangeElements[i];
    var searchResult = rangeElement.getElement().findText(EXPRESSION_REG_EX, null);
    while (searchResult) {
      if (isInsideRangeElement(searchResult, rangeElement)) {
        var match = searchResult.getElement().asText().getText();
        if (searchResult.isPartial()) {
          match = match.substr(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive() - searchResult.getStartOffset() + 1);
        }
        var expression = match.substr(2, match.length - 4).trim();
        var exprObj = { expression: expression, match: match, rangeElement: searchResult };
        Logger.log(exprObj);
        if (expression.substr(0, 3) == 'for' || expression.substr(0, 3) == 'if ') {
          expressionTree = expressionTree.createChild();
          Logger.log('CHILD CREATED WITH THIS EXPRESSION');
        }
        expressionTree.add(exprObj);
        if (expression == 'end') {
          expressionTree = expressionTree.endChild();
        }
      }
      searchResult = rangeElement.getElement().findText(EXPRESSION_REG_EX, searchResult);
    }
  }
  return expressionTree.asArray();
}

function renderExpressions(doc, expressions, input) {
  Logger.log('RENDER EXPRESSIONS');
  Logger.log(expressions);
  for (var i = 0; i < expressions.length; i++) {
    var e = expressions[i];
    if (Array.isArray(e)) {
      renderExpressions(doc, e, input);
    }
    else {
      if (i == 0 && e.expression.substr(0, 4) === 'for ') {
        return renderExpressionFor(doc, expressions, input);
      }
      else if (i == 0 && e.expression.substr(0, 7) === 'for-row') {
        return renderExpressionForRow(doc, expressions, input);
      }
      else if (i == 0 && e.expression.substr(0, 3) === 'if ') {
        return renderExpressionIf(doc, expressions, input);
      }
      else {
        var value = evaluateExpression(e.expression, input);
        Logger.log('REPLACING "' + e.match + '" WITH "' + value + '"');
        var escaped = escapeRegExp(e.match);
        e.rangeElement.getElement().asText().replaceText(escaped, value);
      }
    }
  }
}

function evaluateExpression(expression, input) {
  // split by '|' but jump over quotes
  var exprParts = expression.split(/\|(?=(?:[^'"”]|[”"'][^'"”]*[”"'])*$)/g)
  
  var value = eval('input.' + exprParts[0]);
  // run filters
  for (var i = 1; i < exprParts.length; i++) {
    if (exprParts[i].length == 0) {
      continue;
    }
    // split by ':' but jump over quotes
    var filterParts = exprParts[i].trim().split(/:(?=(?:[^'"”]|[”"'][^'"”]*[”"'])*$)/g)
    Logger.log(filterParts);
    value = runFilter(value, filterParts[0], filterParts.slice(1));
  }
  return value;
}

function runFilter(value, name, args) {
  for (var i = 0; i < args.length; i++) {
    if (args[i].length > 0 && (args[i].charAt(0) == '"' || args[i].charAt(0) == '”')) {
      args[i] = args[i].slice(1, -1); // trim " away
    }
  }
  return filters[name](value, args);
}

/************************************************
 * Filters
 ************************************************/

var filters = { };

filters['date'] = function(value, args) {
  if (args.length > 0) {
    var format = args[0];
    if (typeof value == 'object') {
      value = Utilities.formatDate(value, Session.getScriptTimeZone(), format);
      value = value.replace('Sun', 'su');
      value = value.replace('Mon', 'ma');
      value = value.replace('Tue', 'ti');
      value = value.replace('Wed', 'ke');
      value = value.replace('Thu', 'to');
      value = value.replace('Fri', 'pe');
      value = value.replace('Sat', 'la');
    }
  }
  return value;
};

filters['cut'] = function(value, args) {
  if (args.length > 0) {
    var length = parseInt(args[0]);
    if (typeof value == 'string') {
      if (value.length > length) {
        value = value.substr(0, length) + "…";
      }
    }
  }
  return value;
};

function renderExpressionIf(doc, expressions, input) {
  Logger.log('RENDER IF');
  Logger.log(expressions);
  Logger.log(input);
  var rangeBuilder = doc.newRange();
  //var parts = expressions[0].expression.split(' ');
  var ifClause = expressions[0].expression.substr(3);
  
  var start = expressions[0].rangeElement.getElement();
  var end = expressions[expressions.length - 1].rangeElement.getElement();
  rangeBuilder.addElementsBetween(start, end);
  var range = rangeBuilder.build();
  var rangeElements = range.getRangeElements();
  Logger.log(expressions);
  
  //var parent = rangeElements[1].getElement().getParent();
  //var index = parent.getChildIndex(rangeElements[1].getElement());

  if (eval(ifClause)) {
    var newRangeBuilder = doc.newRange();
    newRangeBuilder.addElementsBetween(rangeElements[1].getElement(), rangeElements[rangeElements.length - 2].getElement());
    var newRange = newRangeBuilder.build();
    var newExpressions = findExpressions(newRange);
    renderExpressions(doc, newExpressions, input);

    // remove {{ if }} and {{ end }}
    //removeParentByType(start, DocumentApp.ElementType.PARAGRAPH);
    //removeElement(start);
    removeParentByType(start, DocumentApp.ElementType.PARAGRAPH);
    removeParentByType(end, DocumentApp.ElementType.PARAGRAPH);
  }
  else {
    // remove whole range
    debugElements(rangeElements);
    for (var i = 0; i < rangeElements.length - 1; i++) {
      Logger.log('REMOVING WHOLE ELEMENT');
      removeElement(rangeElements[i].getElement());
    }
    removeParentByType(end, DocumentApp.ElementType.PARAGRAPH);
  }
}

function renderExpressionFor(doc, expressions, input) {
  Logger.log('RENDER FOR LOOP');
  //Logger.log(expressions);
  var rangeBuilder = doc.newRange();
  //var parts = expressions[0].expression.split(' ');
  var parts = expressions[0].expression.split(/ (?=(?:[^'"”]|[”"'][^'"”]*[”"'])*$)/g)
  var forVariable = null;
  if (parts.length > 3) {
    forVariable = parts[3];
  }
  var start = expressions[0].rangeElement.getElement();
  var end = expressions[expressions.length - 1].rangeElement.getElement();
  rangeBuilder.addElementsBetween(start, end);
  var range = rangeBuilder.build();
  var rangeElements = range.getRangeElements();
  
  var copied = [];
  var body = doc.getBody();
  var parent = rangeElements[1].getElement().getParent();
  var index = parent.getChildIndex(rangeElements[1].getElement());


  var inputVar = eval(parts[1]);
  var count = inputVar.length;
  for (var j = 0; j < count; j++) {
    var copies = [];
    for (var i = 1; i < rangeElements.length - 1; i++) {
      var element = rangeElements[i].getElement();
      copies.push(element.copy());
    }

    var newRangeBuilder = doc.newRange();
    // insert copies (if any)
    for (var i = 0; i < copies.length; i++) {
      Logger.log('COPIABLE ELEMENT TYPE');
      Logger.log(copies[i].getType());
      var copyType = copies[i].getType();
      var newElement = null;
      if (copyType == DocumentApp.ElementType.PARAGRAPH) {
        newElement = parent.insertParagraph(index++, copies[i]);
      }
      else if (copyType == DocumentApp.ElementType.TABLE) {
        newElement = parent.insertTable(index++, copies[i]);
      }
      if (newElement) {
        newRangeBuilder.addElement(newElement);
      }
    }
    if (copies.length) {
      var newRange = newRangeBuilder.build();
      var newExpressions = findExpressions(newRange);
      input[forVariable] = eval(parts[1] + '[' + j + ']');
      renderExpressions(doc, newExpressions, input);
    }
  }
  for (var i = 0; i < rangeElements.length; i++) {
    var element = rangeElements[i].getElement();
    if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
      element.removeFromParent();
    }
    else if (element.getType() == DocumentApp.ElementType.TABLE) {
      element.removeFromParent();
    }
    else {
      removeParentByType(element, DocumentApp.ElementType.PARAGRAPH);
    }
  }
}

function renderExpressionForRow(doc, expressions, input) {
  Logger.log('RENDER FOR-ROW LOOP');
  //Logger.log(expressions);
  var rangeBuilder = doc.newRange();
  var start = expressions[0];
  var parts = start.expression.split(' ');
  var forVariable = null;
  if (parts.length > 3) {
    forVariable = parts[3];
  }
  
  var variable = { };
  variable[parts[3]] = eval(parts[1]);
  Logger.log('VARIABLE');
  Logger.log(variable);
  var end = expressions[expressions.length - 1];
  rangeBuilder.addElementsBetween(start.rangeElement.getElement(), start.rangeElement.getStartOffset(),
                                  end.rangeElement.getElement(), end.rangeElement.getEndOffsetInclusive());
  var range = rangeBuilder.build();

  var elements = range.getRangeElements();
  debugElements(elements);

  // make copy of all TABLE_ROWs between
  var rowElements = filterElementsByType(elements, DocumentApp.ElementType.TABLE_ROW);
  if (rowElements.length > 0) {
    var inputVar = eval(parts[1]);
    var count = inputVar.length;
    var table = rowElements[0].getParent();
    var firstIdx = table.getChildIndex(rowElements[0]);
    for (var j = 0; j < count; j++) {
      var copies = [];
      for (var i = 0; i < rowElements.length; i++) {
        copies.push(rowElements[i].copy());
      }

      // insert copies (if any)
      for (var i = 0; i < copies.length; i++) {
        var newTableRow = table.insertTableRow(firstIdx + i, copies[i]);
        var newRangeBuilder = doc.newRange();
        newRangeBuilder.addElement(newTableRow);
        var newRange = newRangeBuilder.build();
        var newExpressions = findExpressions(newRange);
        input[forVariable] = eval(parts[1] + '[' + j + ']');
        renderExpressions(doc, newExpressions, input);
      }
    }
    // remove originals
    for (var i = 0; i < rowElements.length; i++) {
      rowElements[i].removeFromParent();
    }
  }
  // remove rows having for-now and end
  removeParentByType(start.rangeElement.getElement(), DocumentApp.ElementType.TABLE_ROW);
  removeParentByType(end.rangeElement.getElement(), DocumentApp.ElementType.TABLE_ROW);
}

/**
 * Removes element.
 * If element is not removable, finds removable parent and removes it.
 */
function removeElement(element) {
  while (element.getParent() && !element.removeFromParent) {
    element = element.getParent();
  }
  if (element.removeFromParent) {
    element.removeFromParent();
  }
}

function removeParentByType(element, type) {
  var parent = findParentByType(element, type);
  if (parent) {
    parent.removeFromParent();
  }
}

function filterElementsByType(rangeElements, elementType) {
  var result = [];
  for (var i = 0; i < rangeElements.length; i++) {
    var element = rangeElements[i].getElement();
    if (element.getType() == elementType) {
      result.push(element);
    }
  }
  return result;
}

function debugElements(elements) {
  Logger.log('debugElements ---------------------');
  for (var i = 0; i < elements.length; i++) {
    debugElement(elements[i].getElement());
  }
}

function debugElement(element) {
  var parent;
  var parents = [];
  var str = element.getType() + ', ';
  while (parent = element.getParent()) {
    parents.push(parent.getType());
    element = parent;
  }
  str += parents.join(', ');
  Logger.log(str.toLowerCase());
}

function findParentByType(element, elementType) {
  var parent;
  while (parent = element.getParent()) {
    if (parent.getType() == elementType) {
      return parent;
    }
    element = parent;
  }
  return null;
}

/************************************************
 * Util functions
 ************************************************/

/**
 * Escapes a string to be used in regular expressions.
 *
 * @param {String} str String to be escaped.
 * @return {String} Escaped string.
 */
function escapeRegExp(str) {
  return str.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
}

function twoDigits(num) {
  var str = num.toString();
  if (str.length < 2) {
    str = '0' + str;
  }
  return str;
}

function weekDayAsText(day) {
  var days = [
    'ma',
    'ti',
    'ke',
    'to',
    'pe',
    'la',
    'su'
  ];
  return days[day];
}

function monthAsText(month) {
  var months = [
    'Tammikuu',
    'Helmikuu',
    'Maaliskuu',
    'Huhtikuu',
    'Toukokuu',
    'Kesäkuu',
    'Heinäkuu',
    'Elokuu',
    'Syyskuu',
    'Lokakuu',
    'Marraskuu',
    'Joulukuu'
  ];
  return months[month - 1];
}

/**
 * Returns rows that have the same month and year as defined
 * in MONTH and YEAR.
 */
//function getMonthRows(sheet) {
//  return filterSheetRows(sheet, function(date) {
//    return isSelectedMonth(date);
//  });
//}

function filterSheetRows(sheet, filter) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  var dateColumn = 0;

  var rows = [];
  for (var i = 0; i < values.length; i++) {
    if (i == 0) continue; // skip 1st
    if (values[i].length >= dateColumn + 1) {
      var dateValue = values[i][dateColumn];
      if (filter(dateValue)) {
        rows.push(values[i]);
      }
    }
  }
  return rows;
}

function getMonthRowsByWeek(sheet) {
  var days = findSplitDaysInMonth(parseInt(YEAR), parseInt(MONTH));
  var weekRows = [];
  for (var i = 0; i < days.length - 1; i++) {
    var rows = filterSheetRows(sheet, function(date) {
      return (date >= days[i] && date < days[i+1]);
    });
    weekRows.push(rows);
  }
  return weekRows;
}

function isSelectedMonth(date) {
  if (typeof date == 'object') {
    return (date.getFullYear() == parseInt(YEAR) && date.getMonth() + 1 == parseInt(MONTH));
  }
  return false;
}

function heppa() {
  findSplitDaysInMonth(2016, 4);
}

function findSplitDaysInMonth(year, month) {
  var days = findMondays(year, month);
  // always add monday from next month
  days.push(getNextMonday(days[days.length - 1]));
  return days;
}

function findMondays(year, month) {
  var mondays = [];
  var d = 1;
  var date = new Date(year, month - 1, d++, 0, 0, 0);
  // Hack for getting previous week
  // mondays.push(getPreviousMonday(date));
  while (date.getMonth() == month - 1) {
    if (date.getDay() == 1) {
      mondays.push(date);
    }
    date = new Date(year, month - 1, d++, 0, 0, 0);
  }
  return mondays;
}

function getPreviousMonday(date) {
  var year = date.getFullYear();
  var month = date.getMonth();
  var dateNumber = date.getDate() - 1;
  while (date = new Date(year, month, dateNumber, 0, 0, 0)) {
    if (date.getDay() == 1) {
      return date;
    }
    dateNumber--;
  }
  return null;
}

function getNextMonday(date) {
  var year = date.getFullYear();
  var month = date.getMonth();
  var dateNumber = date.getDate() + 1;
  while (date = new Date(year, month, dateNumber, 0, 0, 0)) {
    if (date.getDay() == 1) {
      return date;
    }
    dateNumber++;
  }
  return null;
}

/**
 * Creates a copy from document.
 *
 * @return {Document} New document.
 */
function copyDocument(sourceId, name) {
  var source = DriveApp.getFileById(sourceId);
  var targetFolder = DriveApp.getFolderById(TARGET_FOLDER);
  var newFile = source.makeCopy(name, targetFolder);
  return DocumentApp.openById(newFile.getId());
}

