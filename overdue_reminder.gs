function checkOverdueUsers() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Paid Users');
    
    if (!sheet) {
      Logger.log('Sheet "Paid Users" not found!');
      return;
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('No data found in "Paid Users" sheet!');
      return;
    }
    
    var headers = data[0];
    var dueDateCol = headers.indexOf('Due Date');
    var nameCol = headers.indexOf('Name');
    var emailCol = headers.indexOf('Email');
    
    if (dueDateCol === -1 || nameCol === -1 || emailCol === -1) {
      Logger.log('Required columns not found!');
      return;
    }
    
    var today = new Date();
    today.setHours(0, 0, 0, 0); // Set to start of day for accurate comparison
    
    var overdueUsers = [];
    
    // Check each row (skip header row)
    for (var i = 1; i < data.length; i++) {
      var dueDateValue = data[i][dueDateCol];
      var name = data[i][nameCol];
      var email = data[i][emailCol];
      
      // Skip if name or email is empty
      if (!name || !email || name === '' || email === '') {
        continue;
      }
      
      // Check if due date is valid and has passed
      if (dueDateValue instanceof Date) {
        var dueDate = new Date(dueDateValue);
        dueDate.setHours(0, 0, 0, 0); // Set to start of day for accurate comparison
        
        if (dueDate < today) {
          overdueUsers.push({
            name: name,
            email: email
          });
        }
      }
    }
    
    // Send email if there are overdue users
    if (overdueUsers.length > 0) {
      sendOverdueEmail(overdueUsers);
    } else {
      Logger.log('No overdue users found.');
    }
    
  } catch (error) {
    Logger.log('Error checking overdue users: ' + error.toString());
  }
}

function sendOverdueEmail(overdueUsers) {
  try {
    var recipient = 'p.kamel@digikala.com';
    var subject = 'Overdue Figma Users';
    
    // Build email body with list of overdue users
    var body = 'The following users have overdue payment dates:\n\n';
    
    for (var i = 0; i < overdueUsers.length; i++) {
      body += overdueUsers[i].name + ' - ' + overdueUsers[i].email + '\n';
    }
    
    // Send email
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });
    
    Logger.log('Overdue email sent successfully to ' + recipient + ' for ' + overdueUsers.length + ' users.');
    
  } catch (error) {
    Logger.log('Error sending overdue email: ' + error.toString());
  }
}

function createDailyTrigger() {
  // Delete existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkOverdueUsers') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new daily trigger at 10:00 AM
  ScriptApp.newTrigger('checkOverdueUsers')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
  
  Logger.log('Daily trigger created successfully for 10:00 AM');
}

