/**
 * Sends a monthly Figma ticket reminder email
 * Runs automatically on the 20th of every month via time-based trigger
 * Can also be triggered manually from the menu
 */
function sendTicketReminder() {
  try {
    // Get current month name
    var now = new Date();
    var monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                     'July', 'August', 'September', 'October', 'November', 'December'];
    var currentMonth = monthNames[now.getMonth()];
    
    // Email details
    var recipient = 'p.kamel@digikala.com';
    var subject = '🔔 Figma Ticket Reminder - ' + currentMonth;
    var body = 'Hi Pooya,\nDon\'t forget this month Figma ticket';
    
    // Send email
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });
    
    // Log success (optional, for debugging)
    Logger.log('Ticket reminder email sent successfully to ' + recipient);
    
    return 'Email sent successfully to ' + recipient;
    
  } catch (error) {
    Logger.log('Error sending ticket reminder: ' + error.toString());
    throw error;
  }
}

/**
 * Sets up a time-based trigger to run sendTicketReminder on the 20th of every month
 * This function should be run once to set up the trigger
 */
function createTicketReminderTrigger() {
  // Delete existing triggers for this function (to avoid duplicates)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendTicketReminder') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new monthly trigger for the 20th of each month
  ScriptApp.newTrigger('sendTicketReminder')
    .timeBased()
    .everyMonths(1)
    .onMonthDay(20)
    .atHour(9) // 9 AM
    .create();
  
  Logger.log('Ticket reminder trigger created successfully');
}

