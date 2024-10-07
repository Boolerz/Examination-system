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
    var teacherEmail = formData[5] ? formData[5].trim() : "";       // Teacher's Email in the 6th column (index 5)
  
    Logger.log("Subject: " + subject);
    Logger.log("Admission Number: " + admissionNumber);
    Logger.log("Score: " + score);
    Logger.log("Teacher Email: " + teacherEmail);
  
    if (!subject || !admissionNumber || !score || !teacherEmail) {
      Logger.log("Missing important data (subject, admission number, score, or teacher email).");
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
      
      // If the cell is already filled, notify both the teacher and the sheet owner
      if (existingScore !== "") {
        Logger.log("Score for student " + admissionNumber + " in " + subject + " already exists: " + existingScore);
  
        // Send email to the teacher notifying them that the marks already exist
        var teacherSubject = "Marks Already Exist";
        var teacherMessage = "Dear Teacher,\n\nYou attempted to enter marks for student " + studentName + 
                             " (Admission Number: " + admissionNumber + ") in the subject " + subject + 
                             ". However, marks already exist for this student and cannot be changed.\n\n" +
                             "Existing Marks: " + existingScore;
        
        MailApp.sendEmail(teacherEmail, teacherSubject, teacherMessage);
        Logger.log("Email notification sent to the teacher: " + teacherEmail);
  
        // Send email to the sheet owner notifying them of the attempted mark change
        var ownerEmail = Session.getActiveUser().getEmail();  // The email of the sheet owner
        var ownerSubject = "Attempted Mark Change Detected";
        var ownerMessage = "Dear Sheet Owner,\n\nA teacher attempted to change the marks for student " + studentName + 
                           " (Admission Number: " + admissionNumber + ") in the subject " + subject + ".\n\n" +
                           "Existing Marks: " + existingScore + "\nAttempted New Marks: " + score + "\n\n" +
                           "Please follow up if necessary.";
  
        MailApp.sendEmail(ownerEmail, ownerSubject, ownerMessage);
        Logger.log("Email notification sent to the sheet owner: " + ownerEmail);
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
  