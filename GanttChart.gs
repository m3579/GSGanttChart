/* Gantt Chart Script!
 
 This script makes dependencies in the Gantt chart template possible!
 
 */

/* This function will be called when any cell is edited (think of it as the entry point into the program) */
function onEdit(e) {
  // Get cells that were edited
  var range = e.range;
  
  // The updateTasks method reconfigures all of the dependencies according to
  // the edit
  updateGanttChart(range);
}

/* The method that makes different types of updates to the Gantt chart based on dependencies (e.g. changed dependency,
   new start time, changed duration, etc.) */
function updateGanttChart(range) {
  
  // Get which column was edited, which tells us which field was changed
  var column = range.getColumn();
  Logger.log(column);
  switch (column) {
    // A task ID was edited
    case 1: {
      updateEditedTaskID(range);
      break;
    }
      
    // A task name was edited
    case 2: {
      updateEditedTaskName(range);
      break;
    }
      
    // A start date was edited
    case 3: {
      updateEditedStartDate(range);
      break;
    }
      
    // A duration was edited
    case 4: {
      updateEditedDuration(range);
      break;
    }
    
    // A dependency was edited
    case 5: {
      updateEditedDependencies(range);
    }
  }
}  

/* A user edited a task ID */
function updateEditedTaskID(range) {
  return;
}  

/* A user edited a task name */
function updateEditedTaskName(range) {
  return;
}  

/* A user edited a task's start date */
function updateEditedStartDate(range) {
  var numTasks = getNumberOfTasks();
  
  // Get the range representing the dependencies (the fifth column from row 3 downwards)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange(3, 5, numTasks);
  
  // For each task, call the updateEditedDependencies function so that the dependencies
  // are set according to the new start date
  for (var row = 1; row <= numTasks; row++) {
    updateEditedDependencies(range.getCell(row, 1));   // the column parameter is defined relative to the range, not the spreadsheet, hence the "1"
  }
}  

/* A user edited a task duration */
function updateEditedTaskDuration(range) { 
  return;
}

/* The user edited a dependency ID (which task(s) a particular task is dependent on);
   time to move a bunch of events forward or backward! */
function updateEditedDependencies(range) {
  var newDependencyVal = range.getDisplayValue();
  
  // No dependencies are declared; reset the task to the beginning
  if (newDependencyVal == "") {
    SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(range.getRow(), 3).setValue(1);
    return;
  }
  
  var newDependencies = newDependencyVal.split(", ");
  
  Logger.log("New dependency IDs: " + newDependencyVal);
  
  // Find when the last dependent task ends to schedule this task for right after that
  
  var latestDependencyEnd = 0;
  
  for (var i = 0; i <+ newDependencies.length; i++) {
    newDependencies[i] = +newDependencies[i]; // convert each element into an integer
    
    // Get start date for each task
    var startDate = -1; // default start date (lets us know whether there was a task with this ID)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    // Iterate over rows in task IDs column
    for (var row = 1; row <= getNumberOfTasks(); row++) {
      // ID in task IDs column matches dependency ID
      if (sheet.getRange(row + 2, 1).getValue() == newDependencies[i]) {    // row + 2 because the first two rows are headers
        // Set startDate to the entry in the Start Date column
        startDate = sheet.getRange(row + 2, 3).getValue();
        break;
      }
    }
    
    // There was no event with this ID - exit the function
    if (startDate == -1) {
      logError("No task with ID " + newDependencies[i]);
      return;
    }
    
    Logger.log("Start date of dependency: " + startDate);
    
    // Add that task's duration to the start date to find when it ends
    
    var duration = SpreadsheetApp.getActiveSheet().getRange(newDependencies[i] + 2, 4).getValue();
    
    Logger.log("Duration of dependency: " + duration);
    
    var endOfDependentTask = startDate + duration;
    
    Logger.log("End of dependency: " + endOfDependentTask);
    
    // Set latestDependencyEnd to the end of the latest dependent task
    
    if (endOfDependentTask > latestDependencyEnd) {
      latestDependencyEnd = endOfDependentTask;
    }
  }
  
  // Set start date for current task to that value
  
  Logger.log(latestDependencyEnd);
  
  Logger.log("Latest dependency end: " + latestDependencyEnd);
  
  var currentTaskRow = range.getRow();
  
  SpreadsheetApp.getActiveSheet().getRange(currentTaskRow, 3).setValue(latestDependencyEnd);
}

function getNumberOfTasks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var numTasks = 0;
  var taskIDIsNotBlank = true;
  var row = 3;
  
  while (taskIDIsNotBlank) {
    var taskID = sheet.getRange(row, 1).getValue();
    if (taskID == "") {
      taskIDIsNotBlank = false;
    }
    else {
      numTasks++;
      row++;
    }
  }
  
  return numTasks;
}

/* Log an error in the Errors column */
function logError(message) {
  Logger.log(message);
  SpreadsheetApp.getActiveSheet().getRange(3, 7).setValue(message);
}
