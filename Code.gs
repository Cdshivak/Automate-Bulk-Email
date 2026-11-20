/**
 * 1. SEND CUSTOMIZED EMAILS WITH ACCOUNT LIMIT PROTECTION
 * Sends emails to rows marked as PENDING or blank, using the full name for greetings,
 * a separate subject section, rich text formatting, Google Drive attachments, and unsubscribe links.
 */
function sendCustomEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // =========================================================================
  // 📝 NON-PROGRAMMER SETTINGS (EDIT THIS SECTION EASILY)
  // =========================================================================
  // 1. Enter your Email Subject Line here:
  var emailSubject = "Important Update & Special Resources";

  // 2. To attach a file, paste its Google Drive File ID here. Leave as "" if no attachment.
  // (Tip: The file ID is the long string found in your file's Google Drive URL)
  var googleDriveFileId = "";
  // =========================================================================

  // 🔍 Check ONLY the currently logged-in user account running the script
  var currentUserEmail = Session.getActiveUser().getEmail();
  var dailyLimit = 300; // Default limit for personal Gmail[cite: 1, 2]
  var accountType = "Personal Gmail";[cite: 2]

  // Check if the user email belongs to a custom Google Workspace domain
  if (
    currentUserEmail &&
    !currentUserEmail.endsWith("@gmail.com") &&
    !currentUserEmail.endsWith("@googlemail.com")
  ) {
    dailyLimit = 1200; //[cite: 1, 2]
    accountType = "Google Workspace"; //[cite: 2]
  }

  console.log("Active User Verified: " + currentUserEmail + " (" + accountType + ")"); //[cite: 2]
  console.log("Enforced Daily Limit: " + dailyLimit + " emails."); //[cite: 2]

  // Track daily usage via script properties
  var scriptProperties = PropertiesService.getScriptProperties(); //[cite: 2]
  var todayStr = new Date().toDateString(); //[cite: 2]
  var lastSentDate = scriptProperties.getProperty("LAST_SENT_DATE"); //[cite: 2]
  var sentTodayCount = 0; //[cite: 2]

  if (lastSentDate === todayStr) {
    sentTodayCount = parseInt(scriptProperties.getProperty("SENT_TODAY_COUNT") || "0"); //[cite: 2]
  } else {
    scriptProperties.setProperty("LAST_SENT_DATE", todayStr); //[cite: 2]
    scriptProperties.setProperty("SENT_TODAY_COUNT", "0"); //[cite: 2]
  }

  if (sentTodayCount >= dailyLimit) {
    console.log("⚠️ Daily limit of " + dailyLimit + " reached for " + currentUserEmail + ". Stopping execution."); //[cite: 2]
    return; //[cite: 2]
  }

  var successCount = 0; //[cite: 1, 2]

  // Loop through all rows in your sheet starting from row 2 (index 1)[cite: 1, 2]
  for (var i = 1; i < data.length; i++) {
    if (sentTodayCount >= dailyLimit) {
      console.log("⚠️ Reached daily limit mid-batch. Stopping."); //[cite: 2]
      break; //[cite: 2]
    }

    var recipientEmail = data[i][0]; // Column A: Email Address[cite: 1, 2]
    var fullName = data[i][1];       // Column B: Full Name[cite: 2]
    var currentStatus = data[i][2];  // Column C: Status[cite: 1, 2]

    // Use the Full Name for the greeting (defaults to "there" if Column B is empty)
    var displayName = (fullName && fullName.toString().trim() !== "") ? fullName : "there"; //[cite: 2]

    // Send only if email is present and status is empty or "PENDING"[cite: 1, 2]
    if (recipientEmail != "" && (currentStatus === "" || currentStatus === "PENDING")) {
      // =====================================================================
      // ✉️ CUSTOMIZE YOUR EMAIL CONTENT HERE (NO CODING NEEDED)
      // Guide:
      // • Use <p>...</p> for new paragraphs.
      // • Use <b>...</b> for bold text.
      // • Use <ul><li>...</li></ul> for bullet points.
      // • Use <a href="...">...</a> for links (like Unsubscribe).
      // =====================================================================
      var htmlBody =
        "<p>Hi <b>" + displayName + "</b>,</p>" +

        "<p>We are excited to share our latest updates with you. Here are the key highlights:</p>" +

        "<ul>" +
          "<li><b>Feature One:</b> Faster performance and simpler workflows.</li>" +
          "<li><b>Feature Two:</b> Completely customizable templates.</li>" +
          "<li><b>Feature Three:</b> Seamless integration with your tools.</li>" +
        "</ul>" +

        "<p>You can review the attached document for more details.</p>" +

        "<p>Best regards,<br><b>Your Team Name</b></p>" +

        "<br><hr>" +
        // Unsubscribe link with recipient email automatically plugged in
        "<p style='font-size: 11px; color: #666;'>If you no longer wish to receive these emails, you can " +
        "<a href='https://yourwebsite.com/unsubscribe?email=" + recipientEmail + "'>unsubscribe here</a>.</p>";

      var options = {
        htmlBody: htmlBody
      };

      // Handle Attachment if a File ID is provided
      if (googleDriveFileId !== "") {
        try {
          var file = DriveApp.getFileById(googleDriveFileId);
          options.attachments = [file.getBlob()];
        } catch (e) {
          console.log("Could not attach file for " + recipientEmail + ": " + e.message);
        }
      }

      // Send the email from the user's account and update status in the sheet
      try {
        GmailApp.sendEmail(recipientEmail, emailSubject, "", options);
        sheet.getRange(i + 1, 3).setValue("SENT");

        successCount++;
        sentTodayCount++;
        scriptProperties.setProperty("SENT_TODAY_COUNT", sentTodayCount.toString());

        console.log("Email sent successfully to: " + recipientEmail + " (Sent today: " + sentTodayCount + "/" + dailyLimit + ")");
      } catch (error) {
        sheet.getRange(i + 1, 3).setValue("FAILED");
        console.log("Failed to send to " + recipientEmail + ": " + error.message);
      }
    }
  }

  console.log("Finished sending batch. Total sent: " + successCount);
}

/**
 * 2. CHECK BOUNCED EMAILS
 * Scans your Gmail inbox for delivery failure notifications,
 * updates status to "BOUNCED", and highlights the entire row light red.
 */
function checkBouncedEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Search Gmail for unread or recent delivery failure notifications[cite: 1, 2]
  var threads = GmailApp.search('from:mailer-daemon subject:("Delivery Status Notification" OR "Failure" OR "Returned mail") newer_than:2d'); //[cite: 1, 2]
  var bouncedCount = 0; //[cite: 2]

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages(); //[cite: 2]
    for (var m = 0; m < messages.length; m++) {
      var body = messages[m].getBody(); //[cite: 2]

      // Loop through all rows in your sheet to see if any email matches the bounce text[cite: 1, 2]
      for (var i = 1; i < data.length; i++) { // Starting from row 2 (index 1)[cite: 2]
        var rowEmail = data[i][0]; //[cite: 2]
        var currentStatus = data[i][2]; //[cite: 2]

        // If the sheet has "SENT" and the email address appears in the bounce message body[cite: 1, 2]
        if (rowEmail != "" && currentStatus === "SENT" && body.indexOf(rowEmail) !== -1) {
          var rowIndex = i + 1; // Convert array index to sheet row[cite: 2]

          // Update status to BOUNCED and highlight row light red[cite: 1, 2]
          sheet.getRange(rowIndex, 3).setValue("BOUNCED"); //[cite: 2]
          sheet.getRange(rowIndex, 1, 1, 3).setBackground("#f4c7c7"); // Light red highlight[cite: 1, 2]

          bouncedCount++; //[cite: 2]
          console.log("Marked as BOUNCED and highlighted red: " + rowEmail); //[cite: 2]
        }
      }
    }
    // Mark the bounce notification email as read so it doesn't get processed twice[cite: 2]
    threads[t].markRead(); //[cite: 2]
  }

  console.log("Finished checking inbox. Total bounces updated: " + bouncedCount); //[cite: 2]
}
