function sendTokensAndAppointments() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
  
    var tokens = [];
    var appointments = [];
  
    var availableSlotsPerLocation = 6; // Number of available slots per location
    var locationQueues = {}; // Track the number of applicants per location
  
    var doctorAppointments = {}; // Track the number of appointments per doctor per location
  
    var filledEmails = {}; // Track filled emails and their corresponding token numbers
  
    for (var i = 1; i < data.length; i++) { // Start from index 1 to skip the header row
      var name = data[i][1];
      var email = data[i][2];
      var location = data[i][4];
      var doctor = data[i][5]; // Assuming doctor info is in column 5, adjust the index if needed
      var language = data[i][6]; // Assuming language is in column 6, adjust the index if needed
  
      var token, appointmentTime;
  
      if (filledEmails[email]) {
        // Use the same token for duplicate email
        token = filledEmails[email];
        appointmentTime = "Duplicate";
      } else {
        if (!locationQueues[location]) {
          locationQueues[location] = {
            count: 0,
            tokens: []
          };
        }
  
        if (
          locationQueues[location].count < availableSlotsPerLocation &&
          (!doctorAppointments[location] || !doctorAppointments[location][doctor] || doctorAppointments[location][doctor] < 2)
        ) {
          locationQueues[location].count++;
          token = location + "-" + locationQueues[location].count;
          locationQueues[location].tokens.push(token);
          appointmentTime = calculateAppointmentTime(locationQueues[location].count);
  
          // Increment the appointment count for the doctor in the specific location
          if (!doctorAppointments[location]) {
            doctorAppointments[location] = {};
          }
          if (!doctorAppointments[location][doctor]) {
            doctorAppointments[location][doctor] = 1;
          } else {
            doctorAppointments[location][doctor]++;
          }
        } else {
          token = "-";
          appointmentTime = "Not available";
        }
  
        filledEmails[email] = token; // Store the token number for the email
      }
  
      tokens.push(token); // Add token to the list
      appointments.push(appointmentTime);
  
      sendEmail(name, email, token, appointmentTime, doctor, language);
    }
  }
  
  function calculateAppointmentTime(token) {
    var startTime = new Date();
    startTime.setHours(10, 0, 0); // Start appointments at 10:00 AM
    var appointmentTime = new Date(startTime.getTime() + (token - 1) * 30 * 60 * 1000); // Add 30 minutes for each token
    return appointmentTime;
  }
  
  function sendEmail(name, email, token, appointmentTime, doctor, language) {
    var subject = "Token and Appointment Information";
    var message = "Dear " + name + ",\n\n";
    message += "Your token number is: " + token + "\n";
    message += "Your appointment time is: ";
  
    if (appointmentTime === "Duplicate") {
      message += appointmentTime + "\n";
      message += "You have a duplicate entry. Your appointment time is the same as the original entry.\n";
    } else {
      message += appointmentTime.toLocaleString() + "\n";
      message += "Please arrive on time for your appointment.\n";
    }
  
    // Add doctor and language information to the email
    message += "Doctor: " + doctor + "\n";
     
    // Send the email
    MailApp.sendEmail(email, subject, message);
  }