  ////////////
 // CONFIG //
////////////

/**
 * Get raw value from configuration sheet given a key.
 *
 * @param  key String name of the config key.
 * @return     Range or null of not found.
 */
function getRawValue(key) {
  try {
    var config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
    if (config == null) {
      return null;
    }
    return config.getRange(key);
  } catch(e) {
    return null;
  }
}

/**
 * Get a value from configuration.
 *
 * @param  key String name of config key
 * @return     Object or null if not found.
 */
function getValue(key) {
  var r = getRawValue(key);
  if (r == null) {
    return null;
  }
  return r.getValue();
}

/**
 * Get values from configuration.
 *
 * @param  key String name of config key
 * @return     Array or null if not found.
 */
function getValues(key) {
  var r = getRawValue(key);
  if (r == null) {
    return null;
  }
  return r.getValues()[0];
}

/**
 * Get the current story ID from configuration.
 *
 * @return int ID.
 */
function getCurrentId() {
  return getValue("nextId");
}

/**
 * Increment story ID and return new ID.
 *
 * @return int ID.
 */
function nextId() {
  nextId = getRawValue("nextId");
  var id = nextId.getValue();
  id += 1;
  nextId.setValue(id);
  return id;
}

/**
 * Get names of backlog sheets.
 *
 * @return Array of String backlog sheet names.
 */
function getBacklogs() {
  return getValues("backlogs");
}

/**
 * Is given sheet a backlog.
 *
 * @param  sheet Spreadsheet object.
 * @return       true if sheet is a backlog, false otherwise.
 */
function isBacklog(sheet) {
  var name = sheet.getName();
  var backlogs = getBacklogs();
  for (var i = 0; i < backlogs.length; ++i) {
    if (name == backlogs[i]) {
      return true;
    }
  }
  return false;
}

/**
 * Get theme names.
 *
 * @return Array of String theme names.
 */
function getThemes() {
  return getValues("themes");
}

/**
 * Get all theme colors
 */
function getAllColors() {
  try {
    var config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
    if (config == null) {
      return THEME_COLORS;
    }
    var themesRange = config.getRange("themes");
    var row = themesRange.getRow();
    var themes = config.getRange(row, 2, 1, 5);
    var themeColors = themes.getBackgrounds()[0];
    if (all(themeColors, "white")) {
      return THEME_COLORS;
    }
    return themeColors;
  } catch(ex) {
    return THEME_COLORS;
  }
}

/*
 *
 * @return int spring length in milliseconds.
 */
function getSprintLength() {
  return getSprintLengthDays() * 24 * 60 * 60 * 1000;
}

/**
 * Get sprint length in days.
 *
 * @return int spring length in days.
 */
function getSprintLengthDays() {
  return getValue("sprintLength");
}

/**
 * Get sprint start day.
 *
 * @return String, one of "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"
 */
function getSprintStartDay() {
  return getValue("sprintStartDay");
}


  /////////
 // IDS //
/////////

/**
 * Auto increment ID column
 */
function onAutoincrementId(event) {
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!isBacklog(sheet))
    return;
  // Get the active row
  var row = sheet.getActiveCell().getRowIndex();
  // Check of ID column is empty
  if (sheet.getRange(row, 1).getValue() == "" &&
      // Make sure theme has something in it to ensure that this is a story.
      sheet.getRange(row, 2).getValue() != "") {
    // Set new ID value
    sheet.getRange(row, 1).setValue(nextId());
  }
}

/**
 * Update sprint data based on sprint length
 */
function onAutoUpdateSprintDate(event) {
}

  ///////////
 // COLOR //
///////////

/**
 * Update row colors for the currently active sheet.
 */
function colorCurrentSheet() {
  colorSheet(SpreadsheetApp.getActiveSheet());
}

/**
 * Update row colors for the given sheet.=
 */
function colorSheet(sheet) {
  if (!isBacklog(sheet))
    return;
  var startRow = 2;
  var endRow = sheet.getLastRow();

  for (var r = startRow; r <= endRow; r++) {
    colorRow(sheet, r);
  }
}

/**
 * Get the color for the given sheet from configuration.
 */
function getColor(themeName) {
  var themes = getRawValue("themes");
  var values = themes.getValues();
  var bg = themes.getBackgroundColors();
  for (var i = 0; i < values[0].length; i++) {
    if (values[0][i] === themeName) {
      return bg[0][i];
    }
  }
  return null;
}

/**
 * Update color for the given row number.
 */
function colorRow(sheet, r){
  var numberOfColumns = sheet.getMaxColumns();
  var dataRange = sheet.getRange(r, 1, 1, numberOfColumns);

  var data = dataRange.getValues();
  var row = data[0];

  var color = getColor(row[1]);
  if (color != null) {
    dataRange.setBackgroundColor(color);
  }
  SpreadsheetApp.flush();
}

/**
 * Event trigger after a row is edited and its color might need to be updated.
 */
function onThemeColor(event) {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!isBacklog(sheet))
    return;
  var r = sheet.getActiveCell().getRowIndex();
  if (r >= 2) {
    colorRow(sheet, r);
  }
}

/**
 * Event trigger to update colors in all backlog sheets.
 */
function onThemeColorAll() {
  var backlogs = getBacklogs();
  for (var i = 0; i < backlogs.length; ++i) {
    var name = backlogs[i];
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    colorSheet(ss);
  }
}


  //////////////
 // BURNDOWN //
//////////////

STATUS_MISSED = "Missed";
STATUS_COMPLETED = "Completed";
STATUS_IN_PROGRESS = "In Progress";
STATUS_PLANNED = "Planned";
STATUS_AUTO = "Auto";

/**
 * Check array values
 *
 * @return true if all values in "a" are "v"
 */
function all(a, v) {
  for (var i = 0; i < a.length; ++i) {
    if (a[i] != v) {
      return false;
    }
  }
  return true;
}

/**
 * Convert a String into an integer.
 *
 * @param s  String
 * @return   int, if cannot parse or null, return 0.
 */
function int(s) {
  i = parseInt(s);
  if (isNaN(i) || s == null) {
    return 0;
  } else {
    return i;
  }
}

/**
 * Convert a String into a float.
 *
 * @param  s String
 * @return   float, if connot parse or null, return 0.0.
 */
function float(s) {
  i = parseFloat(s);
  if (isNaN(i) || s == null) {
    return 0.0;
  } else {
    return i;
  }
}

/**
 * Convert a String into a date.
 *
 * @param  s String
 * @return   Date, if connot parse or null, return null.
 */
function date(s) {
  try {
    return Utilities.formatDate(s, "EST", "yyyy/MM/dd");
  } catch(ex) {
    return null;
  }
}

/**
 * Check empty string
 *
 * @param  s String
 * @return   String, if empty, return null.
 */
function str(s) {
  if (s == null) {
    return null;
  }
  s = s.replace(/^\s+|\s+$/g, '');
  if (s.length <= 0) {
    return null;
  }
  return s;
}

/**
 * Create array populated with value.
 *
 * @param  len int, The length of the array.
 * @param  val Object, the value that is used to populate.
 * @return     Array, an array of size len populated with value.
 */
function arr(len, val) {
    var a = new Array(len);
    while (--len >= 0) {
        a[len] = val;
    }
    return a;
}

/**
 * Sprint ctor.
 *
 * @param  sprintName String, the sprint's name.
 * @param  startDate  Date, the sprint's start date.
 * @param  status     String, one of STATUS_MISSED, STATUS_COMPLETED, STATUS_IN_PROGRESS.
 * @return            Sprint, new sprint object.
 */
function Sprint(sprintName, startDate, status) {
  // Create object
  var self = {};
  // These are filled in after adding to a backlog
  self.index = null;
  self.backlog = null;
  // Spirit number (assumes starts with "Sprint ")
  self.sprintNumber = parseInt(sprintName.substring(7));
  // Strint name
  self.name = sprintName;
  // Sprint start date
  self.startDate = startDate;
  // Sprint start date formatted as a string
  if (startDate != null)
    self.startDateStr = date(startDate);
  // Full name
  self.fullname = self.name + " (" + self.startDateStr + ")";
  // Sprint status
  if (status == "")
    self.status = STATUS_PLANNED;
  else
    self.status = status;
  // Completed points
  self.completed = 0;
  // Missed points
  self.missed = 0;
  // In Progress points
  self.inProgress = 0;
  // Planned points
  self.planned = 0;
  // Private stories array.
  var stories = [];
  /**
   * Add story to this sprint.
   *
   * @param  story Story, the story to add.
   */
  self.push = function(story) {
    stories.push(story);
    // Increment correct counter based on type
    if (story.status == STATUS_COMPLETED) {
      self.completed += story.points;
    } else if (story.status == STATUS_MISSED) {
      self.missed += story.points;
    } else if (story.status == STATUS_IN_PROGRESS) {
      self.inProgress += story.points;
    } else {
      self.planned += story.points;
    }
  };
  // Return object
  return self;
}

/**
 * Story ctor.
 *
 * @param  id      int,    the story's ID.
 * @param  points  int,    number of points this story is worth.
 * @param  status  String, one of STATUS_MISSED, STATUS_COMPLETED, STATUS_IN_PROGRESS.
 * @param  release String, the release name this story belongs to
 * @return         Story,  A new story object.
 */
function Story(id, points, status, release) {
  var self = {};
  // Story ID
  self.id = id;
  // Number of points
  self.points = points;
  // Story status
  self.status = status;
  // Story release
  self.release = release;
  // Set after adding to backlog
  self.sprintNumber = null;
  return self;
}

/**
 * Backlog ctor.
 *
 * @return Backlog, new backlog object.
 */
function Backlog() {
  var self = {};
  // Total number of points across all stories/sprints
  self.totalPoints = 0;
  // List for stories not assigned to any sprint
  self.notAssigned = null;
  // Private list of sprints
  var sprints = [];
  // Private current sprint for adding stories
  var currentSprint = null;

  /**
   * Calculate average valocity of the last 3 sprints given a sprint index.
   *
   * @param  index int, sprint index to calculate back from.
   * @return       int, average velocity.
   */
  var getAverageVelocity = function(index) {
    var count = 0;
    var total = 0;
    while (count < 3 && index >= 0) {
      total += sprints[index].completed;
      count += 1;
      index -= 1;
    }
    return total / count;
  }
  /**
   * Add story to this backlog using the current sprint.
   *
   * @param  story Story, story to add to the current sprint.
   */
  self.pushStory = function(story) {
    // Add to current sprint
    currentSprint.push(story);
    // Set sprint number
    story.sprintNumber = currentSprint.sprintNumber;
    // We don't count missed points
    if (story.status != STATUS_MISSED) {
      self.totalPoints += story.points;
    }
  };
  /**
   * Make the current sprint the Unassigned Sprint.
   * All subsequent pushStory calls will add stories to the Unassigned Sprint.
   */
  self.pushNotAssigned = function() {
    // Add "Not Assigned" sprints, all subsequent stories will be added to it.
    self.notAssigned = Sprint("Not Assigned", null, null);
    currentSprint = self.notAssigned;
  };
  /**
   * Add new sprint and make it the current sprint.
   * All subsequent pushStory calls will add stories to this Sprint.
   *
   * @param  sprint Sprint, sprint to add.
   */
  self.pushSprint = function(sprint) {
    // Set backlog back ref and index.
    sprint.backlog = self;
    sprint.index = sprints.length;
    // Add sprint
    sprints.push(sprint);
    currentSprint = sprint;
  }
  /**
   * Get a sprint by index.
   *
   * @param  i int, sprint index.
   * @return   Sprint, the sprint indexed by i, or undefined if not found.
   */
  self.getSprint = function(i) {
    return sprints[i];
  };
  /**
   * Get a sprints iterator.  A sprints iterator supports a "next" method that will return
   * the next sprint in order everytime it is called.  If none are remaining, it will return null.
   *
   * @return   Iterator, the sprints iterator.
   */
  self.sprints = function() {
    // Create iterator object
    var it = {};
    // Set initial index
    var index = -1;
    // Initialize remaning
    var remaining = self.totalPoints;
    // Average for last 3 completed sprints
    var average = null;
    // Latest sprint start date
    var startDate = null;
    /**
     * Get next sprint.
     *
     * @return   Sprint, the next sprint, or null if none are remaining.
     */
    it.next = function() {
      index += 1;
      var sd = startDate;
      if (index < sprints.length) {
        // Get sprint
        var s = sprints[index];
        // Update start date
        if (s.startDate != null) {
          // Save current sprint date for later
          startDate = s.startDate;
        } else {
          // Incerement date.
          startDate.setDate(startDate.getDate() + getSprintLength());
        }
      } else {
        sd = startDate;
        var sprintName = "Sprint " + (index + 1);
        // Increment start date.
        startDate.setTime(startDate.getTime() + getSprintLength());
        // Setting status to "auto" since this was automatically generated.
        s = Sprint(sprintName, startDate, STATUS_AUTO);
      }
      // Create sprint wrapper
      var sw = {};
      // Copy original sprint attirbutes
      for (var attr in s) {
        sw[attr] = s[attr];
      }
      // Calculated fields
      if (sw.status == STATUS_COMPLETED) {
        // Reduce completed points from remaining
        remaining -= sw.completed;
        // Calculate running average
        average = getAverageVelocity(index);
        // Set average
        sw.averageVelocity = average;
        // Set remaining
        sw.remaining = remaining;
      } else {
        // Future sprints are calculated based on the last average velocity.
        remaining -= average;
        // Remaing for the future is just a forcast
        sw.forcast = remaining;
      }
      if (sw.forcast <= 0) {
        return null;
      } else {
        return sw;
      }
    };
    return it;
  };
  return self;
}

/**
 * Get backlog data from the first backlog defined in the backlog list.
 *
 * @return Backlog, backlog data.
 */
function getBacklogData() {
  // Get the backlog sheet
  var backlogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getBacklogs()[0]);
  // Get all the data we are inetersted in as Object[][]
  var data = backlogSheet.getRange(1, 1, backlogSheet.getLastRow(), 9).getValues();
  var backlog = Backlog();
  for (var i = 0; i < data.length; ++i) {
    // Get sprint name
    var name = data[i][2];
    if (name.indexOf("Sprint ") == 0) {
      // Create new sprint and add to backlog
      current = Sprint(name, data[i][3], str(data[i][7]));
      backlog.pushSprint(current);
    } else if (name == "Not Assigned") {
      // Push not assigned sprint
      backlog.pushNotAssigned();
    } else if (!isNaN(parseInt(data[i][0]))) {
      //Create new story and add to backlog
      var story = Story(parseInt(data[i][0]), int(data[i][4]), data[i][7], data[i][6]);
      backlog.pushStory(story);
    }
  }
  // Return backlog data
  return backlog;
}

/**
 * Generate burndown data, and populate the "Data" sheet.  Also, generate charts.
 */
function onGenerateBurndownData() {
  // Get backlog data
  var backlog = getBacklogData();
  // Initialize column headers
  var headers = ["Name",
                 "Date",
                 "Full Name",
                 "Status",
                 "Remaining",
                 "Completed",
                 "Missed",
                 "In Progress",
                 "Planned",
                 "Average Velocity",
                 "Forcast",
                 "Completion Rate",
                 "'80%",
                 "'100%"]
  var data = [headers];
  // Get sprints
  var sprints = backlog.sprints();
  // Process sprints
  for (var sprint = sprints.next(); sprint != null; sprint = sprints.next()) {
    var completionRate = null;
    var eighty = null;
    var hundred = null;
    if (sprint.status == STATUS_COMPLETED) {
      var completionRate = int(float(sprint.completed) / (float(sprint.completed) + float(sprint.missed)) * 100);
      var eighty = 80;
      var hundred = 100;
    }
    var row = [sprint.name,
               sprint.startDateStr,
               sprint.fullname,
               sprint.status,
               int(sprint.remaining),
               int(sprint.completed),
               int(sprint.missed),
               int(sprint.inProgress),
               int(sprint.planned),
               int(sprint.averageVelocity),
               int(sprint.forcast),
               completionRate,
               eighty,
               hundred];
    data.push(row);
  }
  // Get the data sheet
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  if (dataSheet == null) {
    dataSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Data");
    dataSheet.setFrozenRows(1);
  }
  // Clear old data
  dataSheet.clear();
  // Populate new data
  dataSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

/**
 * Create charts
 *
 * @param  dataSheet Sheet, the data sheet to use for data.
 */
function createCharts(dataSheet) {
  // Delete charts
  var charts = dataSheet.getCharts();
  for (var i = 0; i < charts.length; ++i) {
    dataSheet.removeChart(charts[i]);
  }
  // Create burndown chart
  var chart = dataSheet.newChart().asColumnChart();
  chart.setTitle("Burndown");
  chart.addRange(dataSheet.getRange(1, 3, dataSheet.getLastRow(), 1)); // Data!C1:C100
  chart.addRange(dataSheet.getRange(1, 11, dataSheet.getLastRow(), 1)); // Data!K1:K100
  chart.addRange(dataSheet.getRange(1, 5, dataSheet.getLastRow(), 1)); // Data!E1:E100
  chart.addRange(dataSheet.getRange(1, 6, dataSheet.getLastRow(), 1)); // Data!F1:F100
  chart.addRange(dataSheet.getRange(1, 7, dataSheet.getLastRow(), 1)); // Data!G1:G100
  chart.addRange(dataSheet.getRange(1, 8, dataSheet.getLastRow(), 1)); // Data!H1:H100
  chart.setColors(["#999999", "#3c78d8", "#38761d", "#a61c00", "#b6d7a8"]);
  chart.setStacked();
  chart.setXAxisTitle("Sprints");
  chart.setYAxisTitle("Points");
  chart.setPosition(1, 1, 0, 0);
  dataSheet.insertChart(chart.build());
  // Create predictability chart
  chart = dataSheet.newChart().asLineChart();
  chart.setTitle("Predictability");
  chart.addRange(dataSheet.getRange(1, 3, dataSheet.getLastRow(), 1)); // Data!C1:C100
  chart.addRange(dataSheet.getRange(1, 12, dataSheet.getLastRow(), 3)); // Data!L1:N100
  chart.setColors(["#3c78d8", "#a61c00", "#38761d"]);
  chart.setXAxisTitle("Sprints");
  chart.setYAxisTitle("% Complete");
  chart.setPosition(14, 1, 0, 0);
  chart.setCurveStyle(Charts.CurveStyle.SMOOTH);
  dataSheet.insertChart(chart.build());
}

function onCreateCharts() {
  createCharts(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data"));
}

  ////////////////////
 // INITIALIZATION //
////////////////////

WINDOW_HEIGHT = 400;
WINDOW_WIDTH = 700;
PANEL_HEIGHT = WINDOW_HEIGHT - 60;
WIDGET_WIDTH = "200";
COLUMN_COLOR = "#e0edfe";
PANEL_SPACING = 5;
LIST_BOX_HEIGHT = 15;
POINTS_PER_INCH = 72;
PAGE_MARGIN = 15;
SPRINT_LENGTHS = {
  "1 week":  {"days": 7,  "index": 0},
  "2 weeks": {"days": 14, "index": 1},
  "3 weeks": {"days": 21, "index": 2},
  "4 weeks": {"days": 28, "index": 3}};
DAY_OF_WEEK = {
 "Sunday":    {"name": "Sun", "index": 0},
 "Monday":    {"name": "Mon", "index": 1},
 "Tuesday":   {"name": "Tue", "index": 2},
 "Wednesday": {"name": "Wed", "index": 3},
 "Thursday":  {"name": "Thu", "index": 4},
 "Friday":    {"name": "Fri", "index": 5},
 "Saturday":  {"name": "Sat", "index": 6}};

var HEADERS =       [["ID", "Theme", "Story", "Completion Criteria", "Pt", "Comment", "Release", "Status"]];
var COLUMN_WIDTHS =  [30  , 80     , 500    , 250                 , 30  , 100      , 80       , 110 ];

var THEME_COLORS = ["#9fc5e8", "#b4a7d6", "#ffe599", "#ea9999", "#93c47d"];


/**
 * Get themes from configuration as one string.
 *
 * @return String, theme names, one theme per line.
 */
function getConfigThemes() {
  var themes = getThemes();
  if (themes == null) {
    return "Theme 1\nTheme 2\n";
  }
  return themes.join("\n");
}

/**
 * Get backlog names as one string
 *
 * @return String, backlog names, one backlog per line.
 */
function getConfigBacklogs() {
  var backlogs = getBacklogs();
  if (backlogs == null) {
    return "Backlogs";
  }
  return backlogs.join("\n");
}

/**
 * Get sprint length index from configuration.
 *
 * @return int, sprint length index, or zero if none is found.
 */
function getConfigSprintLengthIndex() {
  var len = getSprintLengthDays();
  for (k in SPRINT_LENGTHS) {
    if (SPRINT_LENGTHS[k]["days"] == len) {
      return SPRINT_LENGTHS[k]["index"];
    }
  }
  return 0;
}

/**
 * Get current story ID from configuration.
 *
 * @return int, current story UID.
 */
function getConfigCurrentId() {
  var currentId = getCurrentId();
  if (currentId == null) {
    return 1;
  }
  return currentId;
}

function getConfigSprintStartDayIndex() {
  var day = getSprintStartDay();
  for (k in DAY_OF_WEEK) {
    if (DAY_OF_WEEK[k]["name"] == day) {
      return DAY_OF_WEEK[k]["index"];
    }
  }
  return 0;
}

/**
 * Create a list box.
 *
 * @param app           UiInstance, The app object that should contain this widget.
 * @param name          String,     The widget's name.
 * @param items         String[],   An array if items.
 * @param selectedIndex int,        The selected item in the list.
 * @return              ListBox     The created list box.
 */
function createListBox(app, name, items, selectedIndex) {
  var list = app.createListBox().setName(name).setWidth(WIDGET_WIDTH); //.setVisibleItemCount(items.length);
  for (var item in items) {
    list.addItem(item);
  }
  list.setSelectedIndex(selectedIndex);
  return list;
}

/**
 * Show create project dialog.
 */
function showCreateProjectDialog() {
  var app = UiApp.createApplication().setTitle("Create Project").setWidth(WINDOW_WIDTH).setHeight(WINDOW_HEIGHT);
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var mainPanel = app.createHorizontalPanel();
  app.add(mainPanel);
  /// Themes Panel
  var themesPanel = app.createVerticalPanel().setSpacing(PANEL_SPACING);
  mainPanel.add(themesPanel);
  themesPanel.add(app.createLabel("Themes").setStyleAttribute("font-weight", "bold"));
  themesPanel.add(app.createTextArea().setName("themeNamesTxt").setWidth(WIDGET_WIDTH).setHeight(PANEL_HEIGHT).setText(getConfigThemes()));
  /// Backlogs Panel
  var backlogsPanel = app.createVerticalPanel().setSpacing(PANEL_SPACING).setStyleAttribute("background-color", COLUMN_COLOR);
  mainPanel.add(backlogsPanel);
  backlogsPanel.add(app.createLabel("Backlogs").setStyleAttribute("font-weight", "bold"));
  backlogsPanel.add(app.createTextArea().setName("backlogNamesTxt").setWidth(WIDGET_WIDTH).setHeight(PANEL_HEIGHT).setText(getConfigBacklogs()));
  /// Misc Panel
  var miscPanel = app.createVerticalPanel().setSpacing(PANEL_SPACING);
  mainPanel.add(miscPanel);
  miscPanel.add(app.createLabel("Sprint Length").setStyleAttribute("font-weight", "bold"));
  miscPanel.add(createListBox(app, "sprintLengthLst", SPRINT_LENGTHS, getConfigSprintLengthIndex()));
  miscPanel.add(app.createLabel("Sprint Start Day").setStyleAttribute("font-weight", "bold"));
  miscPanel.add(createListBox(app, "sprintStartDayLst", DAY_OF_WEEK, getConfigSprintStartDayIndex()));
  miscPanel.add(app.createLabel("Next ID").setStyleAttribute("font-weight", "bold"));
  miscPanel.add(app.createTextBox().setName("nextIdTxt").setWidth(WIDGET_WIDTH).setText(getConfigCurrentId()));
  /// Button Panel
  var buttonPanel = app.createHorizontalPanel().setHorizontalAlignment(UiApp.HorizontalAlignment.RIGHT);
  app.add(buttonPanel);
  buttonPanel.add(app.createSubmitButton("Cancel").setId("submitBtn").addClickHandler(
    app.createServerClickHandler('onCancelBtnClickHandler')));
  buttonPanel.add(app.createSubmitButton("Go").setId("cancelBtn").addClickHandler(
    app.createServerClickHandler('onSubmitBtnClickHandler')
    .addCallbackElement(mainPanel)));
  doc.show(app);
}

/**
 * Set the given values in the supplied spreadsheet, and create a named range.
 *
 * @param ss Sheet, The spreadsheet to use.
 * @param x  int,          Start point to insert data (left most cell).
 * @param y  int,          Start point to insert data (up most cell).
 * @param name String,     Named range name.
 * @param data String[][], The data.
 */
function setValues(ss, x, y, name, data) {
  ss.getRange(x, y, data.length, data[0].length).setValues(data);
  SpreadsheetApp.getActiveSpreadsheet().setNamedRange(name, ss.getRange(x, y + 1, data.length, data[0].length - 1));
}

/**
 * Update config sheet, create if needed.
 *
 * @param themeNames   String[], theme names.
 * @param backlogNames String[], backlog names.
 * @param sprintLength String,   sprint length name.
 * @param nextId       int,      next story ID.
 */
function createConfig(themeNames, backlogNames, sprintLength, sprintStartDay, nextId, themeColors) {
  // Get config sheet.
  var config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (config == null) {
    // Create if needed.
    config = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Config");
  }
  config.clear();

  Logger.log("themeColors: " + themeColors);
  var row = 0
  setValues(config, ++row, 1, "nextId",       [['Next ID', nextId]]);
  setValues(config, ++row, 1, "themes",       [['Themes'].concat(themeNames)]);
  // Set theme colors
  config.getRange(row, 2, 1, themeColors.length).setBackgroundColors([themeColors]);
  setValues(config, ++row, 1, "backlogs",     [['Backlog Sheets'].concat(backlogNames)]);
  setValues(config, ++row, 1, "sprintLength", [['Sprint Length (days)', SPRINT_LENGTHS[sprintLength]["days"]]]);
  setValues(config, ++row, 1, "sprintStartDay", [['Sprint Start Day', DAY_OF_WEEK[sprintStartDay]["name"]]]);
}

/**
 * Create backlog sheets, if they are not present.
 *
 * @param backlogNames String[], backlog names.
 */
function createBacklogs(backlogNames) {
  var backlogs = [];
  // Get spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < backlogNames.length; ++i) {
    // Get backlog sheet.
    var backlog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(backlogNames[i]);
    if (backlog == null) {
      // If backlog does not exist, create it.
      backlog = SpreadsheetApp.getActiveSpreadsheet().insertSheet(backlogNames[i], i++);
      backlog.deleteColumns(HEADERS.length, backlog.getMaxColumns() - HEADERS.length);
      // Create headers.
      var r = backlog.getRange(1, 1, HEADERS.length, HEADERS[0].length)
      r.setValues(HEADERS);
      r.setBackgroundColor("#000");
      r.setFontColor("#FFF");
      r.setFontSize(14);
      r.setFontWeight("bold");
      for (var c = 0; c < COLUMN_WIDTHS.length; ++c) {
        backlog.setColumnWidth(c + 1, COLUMN_WIDTHS[c]);
      }
      ss.setFrozenRows(1);
      backlogs.push(backlog);
    }
  }
  return backlogs;
}

/**
 * Create sample sprints.
 *
 * @param backlogNames Sheet[], backlog names.
 * @param sprintStartDay String, sprint start day of the week (e.g. "Sunday").
 * @param sprintLength String, sprint length name (e.g. "1 week").
 * @param themes       String[], theme names.
 */
function createSprints(backlogs, sprintStartDay, sprintLength, themes) {
  // Set first sprint to the next start of day
  var sprintStartDate1 = new Date();
  var delta = DAY_OF_WEEK[sprintStartDay]["index"] - sprintStartDate1.getDay();
  if (delta < 0) {
    delta = 7 + delta;
  }
  sprintStartDate1.setDate(sprintStartDate1.getDate() + delta);
  // Set second sprint based on first sprint and sprint length.
  var sprintStartDate2 = new Date(sprintStartDate1.getFullYear(), sprintStartDate1.getMonth(), sprintStartDate1.getDate() + SPRINT_LENGTHS[sprintLength]["days"]);
  var sampleSprints = [["", "", "Sprint 1", date(sprintStartDate1)]]
  sampleSprints.push([nextId(), themes[0], "As a <role>, I want <goal/desire> so that <benefit>", ""]);
  sampleSprints.push(["", "", "", ""]);
  sampleSprints.push(["", "", "Sprint 2", date(sprintStartDate2)]);
  sampleSprints.push(["", "", "", ""]);
  sampleSprints.push(["", "", "Unassigned", ""]);

  for (var i = 0; i < backlogs.length; ++i) {
    var backlog = backlogs[i];
    // Create sprints.
    var r = backlog.getRange(2, 1, sampleSprints.length, sampleSprints[0].length)
    r.setValues(sampleSprints);
    // Set fonts and colors
    var rs = [backlog.getRange(2, 1, 1, backlog.getLastColumn()),
              backlog.getRange(5, 1, 1, backlog.getLastColumn()),
              backlog.getRange(7, 1, 1, backlog.getLastColumn())]
    for (var j = 0; j < rs.length; ++j) {
      rs[j].setBackgroundColor("#000");
      rs[j].setFontColor("#FFF");
      rs[j].setFontSize(12);
      rs[j].setFontWeight("bold");
    }
    colorRow(backlog, 3);
  }
}

/**
 * Create validation rules for themes and status.
 *
 * @param backlogs Sheet[], backlog names.
 */
function createValidationRules(backlogs) {
  for (var i = 0; i < backlogs.length; ++i) {
    var backlog = backlogs[i];
    // Validate themes
    var r = backlog.getRange(2, 2, backlog.getMaxRows() - 2, 1);
    var dv = r.getDataValidation();
    dv.requireValuesInRange(SpreadsheetApp.getActive().getRangeByName("themes"));
    dv.setHelpText("Select theme");
    dv.setShowDropDown(true);
    r.setDataValidation(dv);
    // Validate status
    var r = backlog.getRange(2, 8, backlog.getMaxRows() - 2, 1);
    var dv =  r.getDataValidation();
    dv.requireValuesInList([STATUS_IN_PROGRESS, STATUS_COMPLETED, STATUS_MISSED]);
    dv.setHelpText("Select status");
    dv.setShowDropDown(true);
    r.setDataValidation(dv);
    // Color status
    // TODO
  }
}

/**
 * Submit button click handler.
 *
 * @param event Event
 */
function onSubmitBtnClickHandler(event) {
  var app = UiApp.getActiveApplication();

  var themeNames     = event.parameter.themeNamesTxt.trim().split("\n");
  var backlogNames   = event.parameter.backlogNamesTxt.trim().split("\n");
  var sprintLength   = event.parameter.sprintLengthLst;
  var sprintStartDay = event.parameter.sprintStartDayLst;
  var nextId         = event.parameter.nextIdTxt.trim();
  var themeColors    = getAllColors();

  createConfig(themeNames, backlogNames, sprintLength, sprintStartDay, nextId, themeColors);
  var backlogs = createBacklogs(backlogNames);
  createValidationRules(backlogs);
  createSprints(backlogs, sprintStartDay, sprintLength, themeNames);

  onGenerateBurndownData();
  onCreateCharts();

  return app.close();
}

/**
 * Cancel button click handler.
 *
 * Close dialog; do nothing.
 *
 * @param event Event
 */
function onCancelBtnClickHandler(event) {
  var app = UiApp.getActiveApplication();
  return app.close();
}

  //////////////
 // TRIGGERS //
//////////////

function onInstall() {
  // Initiate create project dialog
  showCreateProjectDialog();
}

function onOpen() {
  // Create menus
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Update Burndown", functionName: "onGenerateBurndownData"},
                     {name: "Update Colors", functionName: "colorCurrentSheet"},
                     {name: "Reconfigure Project", functionName: "showCreateProjectDialog"}];
  ss.addMenu("Scrum", menuEntries);
  // Update burndown graph data
  onGenerateBurndownData();
  // Update colors for all rows
  onThemeColorAll();
}

function onEdit() {
  // Update row ID
  onAutoincrementId();
  // Update row color
  onThemeColor();
}
