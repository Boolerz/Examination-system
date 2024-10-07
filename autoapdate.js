function onChange(e) {
    var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
    var broadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Broad Sheet");
  
    if (!formSheet || !broadSheet) {
      Logger.log("Form Responses 1 or Broad Sheet not found.");
      return;
    }
  
    Logger.log("Form submission captured.");
  
    var lastRow = formSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("No data in form responses sheet.");
      return;
    }
  
    var formData = formSheet.getRange(lastRow, 1, 1, formSheet.getLastColumn()).getValues()[0];
    Logger.log("Form data retrieved: " + JSON.stringify(formData));
  
    var subject = formData[1] ? formData[1].trim() : "";            // Subject is in the 2nd column (index 1)
    var studentName = formData[2] ? formData[2].trim() : "";        // Student Name is in the 3rd column (index 2)
    var admissionNumber = formData[3] ? formData[3].toString().trim() : "";  // Admission Number is in the 4th column (index 3)
    var score = formData[4] ? formData[4] : "";                     // Score is in the 5th column (index 4)
    var emailAddress = formData[5] ? formData[5].trim() : "";       // Assuming 6th column is for email address
  
    Logger.log("Subject: " + subject);
    Logger.log("Admission Number: " + admissionNumber);
    Logger.log("Score: " + score);
  
    if (!subject || !admissionNumber || !score) {
      Logger.log("Missing important data (subject, admission number, or score).");
      return;
    }
  
    // Find the correct row in the Broad Sheet using Admission Number
    var admNumberColumn = 1; // Column A in Broad Sheet is for Admission Numbers
    var studentRow = findStudentRow(admissionNumber, broadSheet, admNumberColumn);
  
    Logger.log("Student Row: " + studentRow);
  
    // Find the correct subject column
    var subjectColumn = getSubjectColumn(subject);
  
    Logger.log("Subject Column: " + subjectColumn);
  
    // Check if marks already exist before writing new marks
    if (studentRow !== -1 && subjectColumn !== -1) {
      var existingScore = broadSheet.getRange(studentRow, subjectColumn).getValue();
      
      // If the cell is already filled, send an email notification and don't overwrite
      if (existingScore !== "") {
        Logger.log("Score for student " + admissionNumber + " in " + subject + " already exists: " + existingScore);
        sendEmailNotification(emailAddress, admissionNumber, subject, existingScore);
      } else {
        // If no marks exist, write the new score
        broadSheet.getRange(studentRow, subjectColumn).setValue(score);
        Logger.log("New score added for student " + admissionNumber + ": " + score);
      }
    } else {
      Logger.log("Student Row or Subject Column not found.");
    }
  }
  
  function findStudentRow(admissionNumber, sheet, column) {
    var range = sheet.getRange(2, column, sheet.getLastRow() - 1);  // Skip the header
    var values = range.getValues();
  
    Logger.log("Searching for admission number: " + admissionNumber);
  
    for (var i = 0; i < values.length; i++) {
      Logger.log("Checking row " + (i + 2) + ": " + values[i][0]);
      if (values[i][0].toString().trim() === admissionNumber.trim()) {
        return i + 2;  // Return the correct row, considering header
      }
    }
    return -1; // Return -1 if not found
  }
  
  function getSubjectColumn(subject) {
    var subjectMap = {
      "Maths": 4,      // Maths in Column C
      "Eng": 5,        // English in Column D
      "Kisw": 6,       // Kisw in Column E
      "Chem": 7,       // Chem in Column F
      "Phy": 8,        // Phy in Column G
      "Bio": 9         // Bio in Column H
    };
  
    Logger.log("Looking for subject: " + subject);
    return subjectMap[subject] || -1; // Return -1 if subject not found
  }
  
  function sendEmailNotification(email, admissionNumber, subject, existingScore) {
    var subjectLine = "Score Submission Alert";
    var message = "The score for student " + admissionNumber + " in " + subject + " already exists: " + existingScore + ".";
    
    try {
      MailApp.sendEmail(email, subjectLine, message);
      Logger.log("Notification sent to: " + email);
    } catch (error) {
      Logger.log("Failed to send email: " + error.message);
    }
  }
  