function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var userEmail = Session.getActiveUser().getEmail();
  var memberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Member');
  var data = memberSheet.getRange('A2:D').getValues();
  var userRole = '';

  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === userEmail) {
      userRole = data[i][1];
      break;
    }
  }

  var menu = ui.createMenu('Task Manager')
    .addItem('Update Status', 'showChangeStatusDialog')
    .addItem('Update Progress Report', 'showUpdateProgressDialog')
    .addItem('Search Tasks', 'showSearchDialog')
    .addItem('Filter Tasks', 'showFilterDialog')
    .addItem('Restore View', 'restoreRows')
    .addSeparator();

  if (userRole === 'Employer') {
    menu.addItem('Update Form Options', 'updateFormOptions')
        .addItem('View Done Tasks', 'showViewDoneTasksDialog');
  }

  menu.addToUi();

  // Update task priorities on open
  updateTaskPriorities();
}

function showChangeStatusDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ChangeStatusDialog')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Change Task Status');
}

function showUpdateProgressDialog() {
  var html = HtmlService.createHtmlOutputFromFile('UpdateProgressReportDialog')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Progress Report');
}

function showViewDoneTasksDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ViewDoneTasks')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'View Done Tasks');
}

function showFilterDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FilterTasksDialog')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Filter Tasks');
}

function showSearchDialog() {
  var html = HtmlService.createHtmlOutputFromFile('SearchTasksDialog')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Search Tasks');
}

function changeTaskStatus(taskName, newStatus, progressReport, folderURL) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var userEmail = Session.getActiveUser().getEmail();
  var userRole = getUserRole(userEmail);
  var emailMapping = getEmailMapping();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === taskName) {
      var assignedTo = data[i][7];
      var groupLeader = data[i][8];
      var teamMembers = data[i][9].split(', ');
      var taskType = data[i][6];

      if (userRole === 'Employer' || userEmail === emailMapping[assignedTo] || userEmail === emailMapping[groupLeader]) {
        if (newStatus === 'In Progress' && data[i][3] === 'To Do') {
          data[i][3] = newStatus;
          data[i][4] = progressReport;
          if (taskType === 'Group') {
            logTime(taskName, [groupLeader].concat(teamMembers), 'Start');
          } else {
            logTime(taskName, assignedTo, 'Start');
          }
        } else if (newStatus === 'Done' && data[i][3] === 'In Progress') {
          data[i][3] = newStatus;
          data[i][4] = ''; // Clear progress report
          data[i][5] = folderURL;
          if (taskType === 'Group') {
            logTime(taskName, [groupLeader].concat(teamMembers), 'End');
          } else {
            logTime(taskName, assignedTo, 'End');
          }
          sheet.getRange(i + 1, 1).setBackground('white'); // Remove color when task is done
        } else if (newStatus === 'In Progress' && data[i][3] === 'Done' && userRole === 'Employer') {
          data[i][3] = newStatus;
          if (taskType === 'Group') {
            logTime(taskName, [groupLeader].concat(teamMembers), 'Restart');
          } else {
            logTime(taskName, assignedTo, 'Restart');
          }
          setPriorityColor(sheet.getRange(i + 1, 1), data[i][2]);
        }
        sheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);
        break;
      }
    }
  }
}

function updateProgressReport(taskName, progressReport) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var userEmail = Session.getActiveUser().getEmail();
  var userRole = getUserRole(userEmail);
  var emailMapping = getEmailMapping();
  
  Logger.log('Updating progress report for task: ' + taskName);
  Logger.log('User email: ' + userEmail + ', User role: ' + userRole);

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === taskName && data[i][3] === 'In Progress') {
      var assignedTo = data[i][7];
      var groupLeader = data[i][8];
      var teamMembers = data[i][9].split(', ');

      Logger.log('Found task in progress. Assigned to: ' + assignedTo + ', Group Leader: ' + groupLeader + ', Team Members: ' + teamMembers.join(', '));

      if (userRole === 'Employer' || userEmail === emailMapping[assignedTo] || (userEmail === emailMapping[groupLeader])) {
        Logger.log('User has permission to update the progress report.');
        data[i][4] = progressReport;
        sheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);
        Logger.log('Progress report updated successfully.');
        break;
      } else {
        Logger.log('User does not have permission to update the progress report.');
      }
    }
  }
}

function getEmailMapping() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Member');
  var data = sheet.getRange('A2:D').getValues();
  var emailMapping = {};
  for (var i = 0; i < data.length; i++) {
    emailMapping[data[i][0]] = data[i][2]; // Mapping name to email
  }
  return emailMapping;
}

function rejectTask(taskName, rejectReason) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var emailMapping = getEmailMapping();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === taskName && data[i][3] === 'Done') {
      var assignedTo = data[i][7];
      var groupLeader = data[i][8];
      var teamMembers = data[i][9].split(', ');

      data[i][3] = 'In Progress';
      data[i][5] = ''; // Clear folder URL

      // Log time for rejection
      if (data[i][6] === 'Group') {
        clearEndTime(taskName, [groupLeader].concat(teamMembers));
      } else {
        clearEndTime(taskName, assignedTo);
      }

      // Update the task row
      sheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);
      
      // Reapply color based on due date
      setPriorityColor(sheet.getRange(i + 1, 1), data[i][2]);

      // Send rejection email
      var subject = `Task Rejected: ${taskName}`;
      var message = `Your task "${taskName}" has been rejected for the following reason:\n\n${rejectReason}\n\nPlease address the issues and resubmit the task.`;
      if (emailMapping[assignedTo]) {
        MailApp.sendEmail(emailMapping[assignedTo], subject, message);
      }
      if (emailMapping[groupLeader]) {
        MailApp.sendEmail(emailMapping[groupLeader], subject, message);
      }
      teamMembers.forEach(member => {
        if (emailMapping[member]) {
          MailApp.sendEmail(emailMapping[member], subject, message);
        }
      });

      Logger.log(`Task "${taskName}" rejected and status set to "In Progress" with color reapplied.`);
      break;
    }
  }
}

function clearEndTime(taskName, assignedTo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Time Log');
  if (!sheet) {
    Logger.log('Time Log sheet not found');
    return;
  }

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === taskName && (Array.isArray(assignedTo) ? assignedTo.includes(data[i][1]) : data[i][1] === assignedTo)) {
      sheet.getRange(i + 2, 4, 1, 3).setValues([['', '', '']]);
    }
  }
}

function clearEndTimeInLog(taskName, assignedTo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Time Log');
  var data = sheet.getRange('A2:F').getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === taskName && data[i][1] === assignedTo && data[i][3] !== '') {
      sheet.getRange(i + 2, 4, 1, 3).setValues([['', '', '']]);
      break;
    }
  }
}

function getTaskNames() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var userEmail = Session.getActiveUser().getEmail();
  var userRole = getUserRole(userEmail);
  var taskNames = [];
  var emailMapping = getEmailMapping();

  Logger.log('Fetching tasks for user: ' + userEmail + ' with role: ' + userRole);

  for (var i = 1; i < data.length; i++) {
    var taskName = data[i][0];
    var status = data[i][3];
    var assignedTo = data[i][7];
    var groupLeader = data[i][8];
    var teamMembers = data[i][9] ? data[i][9].split(', ') : [];

    if (status !== 'Done') {
      Logger.log('Checking task: ' + taskName + ', Assigned to: ' + assignedTo + ', Group Leader: ' + groupLeader + ', Team Members: ' + teamMembers.join(', '));
      if (userRole === 'Employer' || userEmail === emailMapping[assignedTo] || userEmail === emailMapping[groupLeader] || teamMembers.includes(emailMapping[userEmail])) {
        taskNames.push(taskName);
      }
    }
  }

  Logger.log('Task names available to user: ' + taskNames.join(', '));
  return taskNames;
}


function getDoneTaskNames() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var taskNames = [];

  for (var i = 1; i < data.length; i++) {
    var taskName = data[i][0];
    var status = data[i][3];

    if (status === 'Done') {
      taskNames.push(taskName);
    }
  }

  return taskNames;
}

function generateTaskReport(taskName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var reportHtml = '<h1>Task Report</h1>';

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === taskName) {
      reportHtml += '<p><strong>Task Name:</strong> ' + data[i][0] + '</p>';
      reportHtml += '<p><strong>Task Description:</strong> ' + data[i][1] + '</p>';
      reportHtml += '<p><strong>Due Date:</strong> ' + data[i][2] + '</p>';
      reportHtml += '<p><strong>Status:</strong> ' + data[i][3] + '</p>';
      reportHtml += '<p><strong>Progress Report:</strong> ' + data[i][4] + '</p>';
      reportHtml += '<p><strong>Folder Name:</strong> ' + data[i][5] + '</p>';
      reportHtml += '<p><strong>Task Type:</strong> ' + data[i][6] + '</p>';
      reportHtml += '<p><strong>Assigned To:</strong> ' + data[i][7] + '</p>';
      reportHtml += '<p><strong>Group Leader:</strong> ' + data[i][8] + '</p>';
      reportHtml += '<p><strong>Team Members:</strong> ' + data[i][9] + '</p>';
      break;
    }
  }
  return reportHtml;
}

function sendRejectionNotification(taskName, assignedTo, rejectionReason) {
  var emailMapping = getEmailMapping();
  var email = emailMapping[assignedTo];

  if (email) {
    var subject = 'Task Rejected: ' + taskName;
    var message = 'Your task has been moved back to "In Progress" with the following rejection reason:\n\n' +
                  rejectionReason + '\n\n' +
                  'Please review and make the necessary adjustments.';
    MailApp.sendEmail(email, subject, message);
  } else {
    Logger.log('Invalid email: ' + assignedTo);
  }
}

function logTime(taskName, assignedTo, action) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Time Log');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Time Log');
    sheet.appendRow(['Task Name', 'Assigned To', 'Start Time', 'End Time', 'Time Spent (hours)', 'Cost']);
  }

  var now = new Date();
  var dateString = Utilities.formatDate(now, 'Asia/Kuala_Lumpur', "dd/MM/yyyy HH:mm:ss");

  if (action === 'Start' || action === 'Restart') {
    if (Array.isArray(assignedTo)) {
      assignedTo.forEach(function(member) {
        sheet.appendRow([taskName, member, dateString, '', '', '']);
      });
    } else {
      sheet.appendRow([taskName, assignedTo, dateString, '', '', '']);
    }
  } else if (action === 'End') {
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === taskName && (Array.isArray(assignedTo) ? assignedTo.includes(data[i][1]) : data[i][1] === assignedTo) && data[i][3] === '') {
        var startTime = new Date(data[i][2]);
        var timeSpent = (now - startTime) / (1000 * 60 * 60);
        var hourlyRate = getHourlyRate(data[i][1]);
        var cost = timeSpent * hourlyRate;
        sheet.getRange(i + 2, 4, 1, 3).setValues([[dateString, timeSpent, cost]]);
      }
    }
  }
}

function getUserRole(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Member');
  var data = sheet.getRange('A2:D').getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === email) {
      Logger.log('User role found: ' + data[i][1]);
      return data[i][1];
    }
  }
  Logger.log('User role not found for email: ' + email);
  return null;
}

function getHourlyRate(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Member');
  var data = sheet.getRange('A2:D').getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === name) {
      return parseFloat(data[i][3]);
    }
  }
  return 0;
}

function updateFormOptions() {
  var formId = '1ElgmUpsM24txcMSLNM134ed-DJ8qOLCGiEily_GM9J0'; // Replace with your form ID
  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('Form opened successfully.');
  } catch (e) {
    Logger.log('Error opening form: ' + e.message);
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Member');
  var data = sheet.getRange('A2:D').getValues();
  var employees = data.filter(row => row[1] === 'Employee').map(row => row[0]);

  // Use trimmed field names
  updateFormField(form, 'Assigned To', employees);
  updateFormField(form, 'Group Leader', employees);
  updateFormField(form, 'Team Members', employees);
}

function updateFormField(form, fieldName, choices) {
  Logger.log('Updating field: ' + fieldName + ' with choices: ' + choices);
  var items;
  try {
    items = form.getItems();
    Logger.log('Form Items: ' + items.map(item => item.getTitle().trim()));
  } catch (e) {
    Logger.log('Error getting form items: ' + e.message);
    return;
  }

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (item.getTitle().trim() === fieldName.trim()) {
      var itemType = item.getType();
      if (itemType === FormApp.ItemType.MULTIPLE_CHOICE || itemType === FormApp.ItemType.LIST) {
        item.asMultipleChoiceItem().setChoiceValues(choices);
      } else if (itemType === FormApp.ItemType.CHECKBOX) {
        item.asCheckboxItem().setChoiceValues(choices);
      }
    }
  }
}

function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form responses 1');
  var lastRow = sheet.getLastRow();
  var rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  Logger.log('Form submission data: ' + rowData);

  var timeStamp = new Date(rowData[0]);
  var taskName = rowData[1];
  var taskType = rowData[2];
  var dueDate = new Date(rowData[3]);
  dueDate.setHours(23);
  dueDate.setMinutes(59);
  var dueDateString = formatDateToMalaysia(dueDate);
  var taskDescription = rowData[4];
  var assignedTo = taskType === 'Individual' ? rowData[5] : '';
  var groupLeader = taskType === 'Group' ? rowData[6] : '';
  var teamMembers = taskType === 'Group' ? rowData[7].split(',').map(member => member.trim()) : [];

  Logger.log('Task details: ' + taskName + ', ' + taskType + ', ' + dueDateString + ', ' + taskDescription + ', ' + assignedTo + ', ' + groupLeader + ', ' + teamMembers);

  addTaskToKanbanBoard(taskName, taskDescription, dueDateString, 'To Do', '', '', taskType, assignedTo, groupLeader, teamMembers.join(', '));
  sendEmailNotification(taskName, taskType, dueDate, taskDescription, assignedTo, groupLeader, teamMembers);
  createCalendarEvent(taskName, timeStamp, dueDate, taskDescription, taskType, assignedTo, groupLeader, teamMembers);
}

function formatDateToMalaysia(date) {
  var formattedDate = Utilities.formatDate(date, 'Asia/Kuala_Lumpur', "EEE MMM dd yyyy HH:mm:ss") + ' (Malaysia Standard Time)';
  return formattedDate;
}

function addTaskToKanbanBoard(taskName, taskDescription, dueDate, status, progressReport, folderName, taskType, assignedTo, groupLeader, teamMembers) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var lastRow = sheet.getLastRow();
  var taskData = [taskName, taskDescription, dueDate, status, progressReport, folderName, taskType, assignedTo, groupLeader, teamMembers];
  Logger.log('Task Data: ' + taskData);
  sheet.getRange(lastRow + 1, 1, 1, taskData.length).setValues([taskData]);
  setPriorityColor(sheet.getRange(lastRow + 1, 1), dueDate);
}

function sendEmailNotification(taskName, taskType, dueDate, taskDescription, assignedTo, groupLeader, teamMembers) {
  var emailMapping = getEmailMapping();

  if (taskType === 'Individual') {
    var email = emailMapping[assignedTo];
    if (email) {
      sendEmail(email, taskName, dueDate, taskDescription);
    } else {
      Logger.log('Invalid email: ' + assignedTo);
    }
  } else if (taskType === 'Group') {
    var emails = [];
    if (emailMapping[groupLeader]) {
      emails.push(emailMapping[groupLeader]);
    }
    teamMembers.forEach(member => {
      if (emailMapping[member]) {
        emails.push(emailMapping[member]);
      } else {
        Logger.log('Invalid email: ' + member);
      }
    });
    emails.forEach(email => {
      sendEmail(email, taskName, dueDate, taskDescription);
    });
  }
}

function sendEmail(email, taskName, dueDate, taskDescription) {
  var subject = 'New Task Assigned: ' + taskName;
  var message = 'You have been assigned a new task.\n\n' +
                'Task Name: ' + taskName + '\n' +
                'Due Date: ' + dueDate + '\n' +
                'Task Description: ' + taskDescription + '\n\n' +
                'Please check the task management system for more details.';

  MailApp.sendEmail(email, subject, message);
}

function createCalendarEvent(taskName, timeStamp, dueDate, taskDescription, taskType, assignedTo, groupLeader, teamMembers) {
  var calendar = CalendarApp.getDefaultCalendar();

  var emails = [];
  var emailMapping = getEmailMapping();
  if (taskType === 'Individual') {
    var email = emailMapping[assignedTo];
    if (email) emails.push(email);
  } else if (taskType === 'Group') {
    if (emailMapping[groupLeader]) emails.push(emailMapping[groupLeader]);
    teamMembers.forEach(member => {
      if (emailMapping[member]) emails.push(emailMapping[member]);
    });
  }

  var event = calendar.createEvent(taskName, new Date(timeStamp), new Date(dueDate), {
    description: taskDescription,
    guests: emails.join(','),
    sendInvites: true
  });

  if (taskType === 'Group') {
    var conference = CalendarApp.newConferenceDataBuilder()
      .setConferenceSolution(CalendarApp.ConferenceSolution.HANGOUTS_MEET)
      .setEntryPoints([
        CalendarApp.newEntryPointBuilder().setEntryPointType(CalendarApp.EntryPointType.VIDEO).build()
      ])
      .build();

    event.setConferenceData(conference);
  }
}

function updateTaskPriorities() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var now = new Date();

  for (var i = 1; i < data.length; i++) {
    var status = data[i][3];
    var cell = sheet.getRange(i + 1, 1); // Task Name cell
    
    if (status === 'Done') {
      cell.setBackground('white'); // Remove color when task is done
    } else {
      var dueDate = new Date(data[i][2]);
      var daysLeft = (dueDate - now) / (1000 * 60 * 60 * 24);

      if (daysLeft < 7) {
        cell.setBackground('red');
      } else if (daysLeft <= 14) {
        cell.setBackground('yellow');
      } else {
        cell.setBackground('green');
      }
    }
  }
}

function setPriorityColor(cell, dueDate) {
  var now = new Date();
  var dueDateObj = new Date(dueDate);
  var daysLeft = (dueDateObj - now) / (1000 * 60 * 60 * 24);

  if (daysLeft < 7) {
    cell.setBackground('red'); // High priority
  } else if (daysLeft <= 14) {
    cell.setBackground('yellow'); // Medium priority
  } else {
    cell.setBackground('green'); // Low priority
  }
}

function checkDueDates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var alertDays = 3; // Number of days before due date to send alert
  var emailNotifications = [];

  Logger.log('Checking due dates...');
  
  for (var col = 1; col <= 2; col++) { // Check columns "To Do" and "In Progress"
    for (var row = 1; row < data.length; row++) {
      if (data[row][col - 1]) {
        var taskDetails = data[row][col - 1].split('\n');
        var dueDate = new Date(taskDetails[5]);
        var daysLeft = (dueDate - now) / (1000 * 60 * 60 * 24);

        Logger.log('Task: ' + taskDetails[0] + ', Due in: ' + daysLeft + ' days');

        if (daysLeft <= alertDays && daysLeft >= 0) {
          var taskName = taskDetails[0];
          var assignedTo = taskDetails[2];
          var email = getEmailForAssignee(assignedTo);
          if (email) {
            emailNotifications.push({ email: email, taskName: taskName, dueDate: dueDate, assignedTo: assignedTo });
            Logger.log('Alert added for task: ' + taskName);
          }
        }
      }
    }
  }

  sendDueDateAlerts(emailNotifications);
}

function sendDueDateAlerts(notifications) {
  for (var i = 0; i < notifications.length; i++) {
    var notification = notifications[i];
    var subject = 'Task Due Soon: ' + notification.taskName;
    var message = 'Dear ' + notification.assignedTo + ',\n\n' +
                  'This is a reminder that the following task is due soon:\n\n' +
                  'Task Name: ' + notification.taskName + '\n' +
                  'Due Date: ' + Utilities.formatDate(notification.dueDate, 'Asia/Kuala_Lumpur', "dd/MM/yyyy") + '\n\n' +
                  'Please ensure that the task is completed on time.\n\n' +
                  'Best regards,\nTask Management System';

    Logger.log('Sending email to: ' + notification.email + ' for task: ' + notification.taskName);
    MailApp.sendEmail(notification.email, subject, message);
  }
}

function filterTasks(filters) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();
  var userEmail = Session.getActiveUser().getEmail();
  var userRole = getUserRole(userEmail);
  var emailMapping = getEmailMapping();
  var userTasks = getTaskNames();

  Logger.log('User email: ' + userEmail);
  Logger.log('Filters applied: ' + JSON.stringify(filters));
  Logger.log('User tasks: ' + JSON.stringify(userTasks));

  for (var i = 1; i < data.length; i++) {
    var showRow = true;
    var taskName = data[i][0];
    var status = data[i][3];
    var color = sheet.getRange(i + 1, 1).getBackground();
    var assignedTo = data[i][7];
    var groupLeader = data[i][8];
    var teamMembers = data[i][9] ? data[i][9].split(', ') : [];

    Logger.log('Task: ' + taskName + ' | Status: ' + status + ' | Color: ' + color + ' | Assigned To: ' + assignedTo + ' | Group Leader: ' + groupLeader + ' | Team Members: ' + teamMembers.join(', '));

    if (filters.taskInvolveMe) {
      var taskInvolvesUser = (emailMapping[assignedTo] === userEmail) || (emailMapping[groupLeader] === userEmail);
      if (!taskInvolvesUser) {
        for (var j = 0; j < teamMembers.length; j++) {
          if (emailMapping[teamMembers[j]] === userEmail) {
            taskInvolvesUser = true;
            break;
          }
        }
      }
      Logger.log('Checking if task involves user. AssignedTo: ' + assignedTo + ', GroupLeader: ' + groupLeader + ', TeamMembers: ' + teamMembers.join(', ') + ', taskInvolvesUser: ' + taskInvolvesUser);
      showRow = showRow && taskInvolvesUser;
    }

    if (filters.status && filters.status.length > 0) {
      var statusMatch = filters.status.includes(status);
      Logger.log('Checking status filter. Status: ' + status + ', StatusMatch: ' + statusMatch);
      showRow = showRow && statusMatch;
    }

    if (filters.priority && filters.priority.length > 0) {
      var priority = '';
      if (color === '#ff0000') {
        priority = 'High';
      } else if (color === '#ffff00') {
        priority = 'Medium';
      } else if (color === '#00ff00') {
        priority = 'Low';
      }
      var priorityMatch = filters.priority.includes(priority);
      Logger.log('Checking priority filter. Priority: ' + priority + ', PriorityMatch: ' + priorityMatch);
      showRow = showRow && priorityMatch;
    }

    Logger.log('Task: ' + taskName + ' | Show Row: ' + showRow);

    if (showRow) {
      sheet.showRows(i + 1);
    } else {
      sheet.hideRows(i + 1);
    }
  }
}

function searchTasks(keyword) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowText = data[i].join(' ').toLowerCase();
    var showRow = rowText.includes(keyword.toLowerCase());

    if (showRow) {
      sheet.showRows(i + 1);
    } else {
      sheet.hideRows(i + 1);
    }
  }
}

function restoreRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kanban Board');
  var lastRow = sheet.getLastRow();
  sheet.showRows(1, lastRow);
}


