/**
 * Invoice Automation using Google Forms + Sheets + Apps Script
 * 
 * This file is intentionally kept as a single script.gs
 * to match Google Apps Script editor structure.
 */

const ROOM_TEMPLATE_FILE_ID = 'PUT_ROOM_TEMPLATE_DOC_ID_HERE';
const FOOD_TEMPLATE_FILE_ID = 'PUT_FOOD_TEMPLATE_DOC_ID_HERE';
const CURRENCY_SIGN = 'Rs. ';
// All generated invoices witll be stored in this folder
const INVOICE_MAIN_FOLDER = "Invoices";

// Converts a float to a string value in the desired currency format
function toCurrency(num) {
    var fmt = Number(num).toFixed(2);
    return `${CURRENCY_SIGN}${fmt}`;
}

// Format datetimes to: YYYY-MM-DD
function toDateFmt(dt_string) {
  var millis = Date.parse(dt_string);
  var date = new Date(millis);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);

  // Return the date in YYYY-mm-dd format
  return `${year}-${month}-${day}`;
}


// Parse and extract the data submitted through the form.
function parseFormData(values, header) {

    // Set temporary variables to hold prices and data.
    // GST rate added just for demo purpose.
    // Please adjust GST as per regulations.
    var rgst = 0.025;
    var fgst = 0.025;
    var rpd = 0;
    var days = 0;
    var persons = 0;
    var response_data = {};
    var food_data = {};
    response_data['generate_food_invoice'] = false;

    // Iterate through all of our response data and add the keys (headers)
    // and values (data) to the response dictionary object.
    // Change this as per your form fields and customization needed.
    for (var i = 0; i < values.length; i++) {
      // Extract the key and value
      var key = header[i];
      var value = values[i];

      if (key.toLowerCase() == "rate/day") {
        rpd = value;
      }
      else if (key.toLowerCase() == "days") {
        days = value;
      }
      else if (key.toLowerCase() == "persons") {
        persons = value;
      }
      else if (key.toLowerCase() == "rate/day 2") {
        rpd2 = value;
      }
      else if (key.toLowerCase() == "days 2") {
        days2 = value;
      }
      else if (key.toLowerCase() == "persons 2") {
        persons2 = value;
      }
      else if (key.toLowerCase() == "rate/day 3") {
        rpd3 = value;
      }
      else if (key.toLowerCase() == "days 3") {
        days3 = value;
      }
      else if (key.toLowerCase() == "persons 3") {
        persons3 = value;
      }
      else if (key.toLowerCase() == "rate/day 4") {
        rpd4 = value;
      }
      else if (key.toLowerCase() == "days 4") {
        days4 = value;
      }
      else if (key.toLowerCase() == "persons 4") {
        persons4 = value;
      }
      else if (key.toLowerCase().includes("check in")) {
        check_in_date = new Date(value);
        value = Utilities.formatDate(check_in_date,'Asia/Kolkata','dd/MM/yyyy');
        check_in_time = Utilities.formatDate(check_in_date,'Asia/Kolkata','hh:mm a');
        response_data['date'] = Utilities.formatDate(check_in_date,'Asia/Kolkata','dd');
      }
      else if (key.toLowerCase().includes("check out")) {
        check_out_date = new Date(value);
        value = Utilities.formatDate(check_out_date,'Asia/Kolkata','dd/MM/yyyy');
        check_out_time = Utilities.formatDate(check_out_date,'Asia/Kolkata','hh:mm a');
      }
      else if (key.toLowerCase().includes("invoice date")) {
        invoice_date = new Date(value);
        value = Utilities.formatDate(invoice_date,'Asia/Kolkata','dd/MM/yyyy');
      }
      else if (key.toLowerCase() === "food1 qty") {
        food_data['food1_qty'] = value;
        if(value !== '')
          response_data['generate_food_invoice'] = true;
      }
      else if (key.toLowerCase() === "food2 qty") {
        food_data['food2_qty'] = value;
      }
      else if (key.toLowerCase() === "food3 qty") {
        food_data['food3_qty'] = value;
      }
      else if (key.toLowerCase() === "food4 qty") {
        food_data['food4_qty'] = value;
      }
      else if (key.toLowerCase() === "food5 qty") {
        food_data['food5_qty'] = value;
      }
      else if (key.toLowerCase() === "food6 qty") {
        food_data['food6_qty'] = value;
      }
      else if (key.toLowerCase() === "food7 qty") {
        food_data['food7_qty'] = value;
      }
      else if (key.toLowerCase() === "food1 rate") {
        food_data['food1_rate'] = value;
      }
      else if (key.toLowerCase() === "food2 rate") {
        food_data['food2_rate'] = value;
      }
      else if (key.toLowerCase() === "food3 rate") {
        food_data['food3_rate'] = value;
      }
      else if (key.toLowerCase() === "food4 rate") {
        food_data['food4_rate'] = value;
      }
      else if (key.toLowerCase() === "food5 rate") {
        food_data['food5_rate'] = value;
      }
      else if (key.toLowerCase() === "food6 rate") {
        food_data['food6_rate'] = value;
      }
      else if (key.toLowerCase() === "food7 rate") {
        food_data['food7_rate'] = value;
      }
  
      // Add the key/value data pair to the response dictionary.
      response_data[key] = value;
    }
    
    response_data["check_in_time"] = check_in_time;
    response_data["check_out_time"] = check_out_time;

    // Once all data is added, adjust the subtotal and total
    var stay_total = rpd * days;
    var stay_total2 = rpd2 * days2;
    var stay_total3 = rpd3 * days3;
    var stay_total4 = rpd4 * days4;

    response_data["stay_total"] = Number(stay_total).toFixed(2);
    response_data["stay_total2"] = Number(stay_total2).toFixed(2);
    response_data["stay_total3"] = Number(stay_total3).toFixed(2);
    response_data["stay_total4"] = Number(stay_total4).toFixed(2);

    if (response_data["stay_total2"] <= 0) {
       response_data["stay_total2"] = ''
    }
    if (response_data["stay_total3"] <= 0) {
       response_data["stay_total3"] = ''
    }
    if (response_data["stay_total4"] <= 0) {
       response_data["stay_total4"] = ''
    }

    var grand_total_wo_gst = stay_total + stay_total2 + stay_total3 + stay_total4;
    var rSgst = grand_total_wo_gst * rgst;
    var rCgst = grand_total_wo_gst * rgst;

    response_data["grand_total"] = Number(grand_total_wo_gst).toFixed(2);
    response_data["sgst"] = Number(rSgst).toFixed(2);
    response_data["cgst"] = Number(rCgst).toFixed(2);

    var net_stay_total = customRound(grand_total_wo_gst + rSgst + rCgst);
    response_data["total"] = toCurrency(net_stay_total);
    response_data['total_in_words'] = NumInWords(net_stay_total);

    // Food details
    if(response_data['generate_food_invoice']) {
      var result = calculateTotalRateAndGrandTotal(food_data);
      for (var i = 1; i <= 7; i++) { // Adjust as per no. of food items.
        var totalKey = 'food' + i + '_total';
        if (result.data[totalKey] !== undefined) {
          response_data[totalKey] = Number(result.data[totalKey]).toFixed(2);
        }else {
          response_data[totalKey] = '';
        }
      }
      var food_total = result.grandTotal;
      response_data['food_total'] = Number(food_total).toFixed(2);

      var fSgst = food_total * fgst;
      var fCgst = food_total * fgst;

      response_data["food_sgst"] = Number(fSgst).toFixed(2);
      response_data["food_cgst"] = Number(fCgst).toFixed(2);

      var food_net_total = customRound(food_total + fSgst + fCgst);

      response_data["food_net_total"] = toCurrency(food_net_total);
      response_data['food_total_in_words'] = NumInWords(food_net_total);

    }

    return response_data;
}

// Helper function to modify/calculate food data.
function calculateTotalRateAndGrandTotal(data) {
  var newData = Object.assign({}, data);
  var grandTotal = 0;

  for (var i = 1; i <= 7; i++) { // Assuming there are 7 food items
    var qtyKey = 'food' + i + '_qty';
    var rateKey = 'food' + i + '_rate';
    var totalKey = 'food' + i + '_total';

    var qty = newData[qtyKey];
    var rate = newData[rateKey];
    
    if (qty !== '' && rate !== '') {
      var totalRate = qty * rate;
      newData[totalKey] = totalRate;
      grandTotal += totalRate;
    }
  }
  return {
    data: newData, // Return the modified data object
    grandTotal: grandTotal // Return the grand total
  };
}

// Helper function to inject data into the template
function populateTemplate(document, response_data) {

    // Get the document header and body (which contains the text we'll be replacing).
    var document_body = document.getBody();
    // Replace variables in the header
    for (var key in response_data) {
      var match_text = `{{${key}}}`;
      var value = response_data[key];
      // Replace our template with the final values
      try{
        document_body.replaceText(match_text, value);
      }catch(e){
        console.log(e, key, value);
      }
    }

}


// Function to populate the template form
function createDocFromForm() {

  // Get active sheet and tab of our response data spreadsheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  var last_row = sheet.getLastRow() - 1;

  // Get the data from the spreadsheet.
  var range = sheet.getDataRange();
 
  // Identify the most recent entry and save the data in a variable.
  var data = range.getValues()[last_row];
  
  // Extract the headers of the response data to automate string replacement in our template.
  var headers = range.getValues()[0];

  // Parse the form data.
  var response_data = parseFormData(data, headers);
  const roomPdfFile = generateInvoice(response_data);
  if(response_data['generate_food_invoice']) {
    var foodPdfFile = generateFoodInvoice(response_data);
  }

  sendEmail(response_data['Customer Email'], roomPdfFile, foodPdfFile, response_data['Customer Email CC']);

}

function generateInvoice(response_data){
  var yearMonthFolder = getYearMonthFolder();
   // Retreive the template file and destination folder.
  var template_file = DriveApp.getFileById(ROOM_TEMPLATE_FILE_ID);
  var target_folder = yearMonthFolder.monthFolder;

  // Copy the template file so we can populate it with our data.
  // Generate name of the file. Generated as per the request you can modify as per your requirement.
  // Format {Date}-Room-{Guest-name}-{Invoice No}.pdf
  var filename = `${response_data["date"]}-Room-${response_data["Guest Name"].replace(/ /g, "-")}-${response_data["Invoice Number"]}`;
  var document_copy = template_file.makeCopy(filename, target_folder);

  // Open the copy.
  var document = DocumentApp.openById(document_copy.getId());

  // Populate the template with our form responses and save the file.
  populateTemplate(document, response_data);

  document.saveAndClose();
  // Export as PDF.
  var BLOBPDF = document.getAs(MimeType.PDF);
  var pdfFile =  target_folder.createFile(BLOBPDF).setName(filename);
  var pdfUrl = pdfFile.getUrl();
  // Delete the Doc file if not needed.(Optional)
  document_copy.setTrashed(true);
  return pdfUrl;
}

function generateFoodInvoice(response_data){
  var yearMonthFolder = getYearMonthFolder();
   // Retreive the template file and destination folder.
  var template_file = DriveApp.getFileById(FOOD_TEMPLATE_FILE_ID);
  var target_folder = yearMonthFolder.monthFolder;

  // Copy the template file so we can populate it with our data.
  // Generate name of the file. Generated as per the request you can modify as per your requirement.
  // Format {Date}-Food-{Guest-name}-B-{Invoice No}.pdf
  var filename = `${response_data["date"]}-Food-${response_data["Guest Name"].replace(/ /g, "-")}-B-${response_data["Invoice Number"]}`;
  var document_copy = template_file.makeCopy(filename, target_folder);

  // Open the copy.
  var document = DocumentApp.openById(document_copy.getId());

  // Populate the template with our form responses and save the file.
  populateTemplate(document, response_data);
  document.saveAndClose();
  // Export as PDF.
  var BLOBPDF = document.getAs(MimeType.PDF);
  var pdfFile =  target_folder.createFile(BLOBPDF).setName(filename);
  var pdfUrl = pdfFile.getUrl();
  // Delete the Doc file if not needed.(Optional)
  document_copy.setTrashed(true);
  return pdfUrl;
}

function sendEmail(customer_email,roomPdfFile, foodPdfFile, customerEmailCC){

  var htmlBody = 'Hi Sir/Madam,';
  if (roomPdfFile) {
    htmlBody += "<p>Please find the invoice soft copy in the links below.</p><br>" +
                 "<strong><a href='" + roomPdfFile + "'>View Room Invoice</a></strong>";
  }

  if (foodPdfFile) {
    htmlBody += "<br><strong><a href='" + foodPdfFile + "'>View Food Invoice</a></strong>";
  }

  htmlBody += '<br><p>Hope you enjoyed your stay with us.<br>Do visit again.</p><br><p>Regards,<br>Your Company Name</p>'
              +'<p>This is an automated service.<br>Please do not reply to this message.</p>';

  if(customer_email) {
    GmailApp.sendEmail(customer_email, "Automated invoice", "Automated invoice customer copy", {
      htmlBody: htmlBody,
      cc: customerEmailCC,  
      name: "Automated invoice"
    });
  }

  GmailApp.sendEmail('YourSystemEmail@gmail.com', "New invoice generated", "Please find the attachment below", {
    htmlBody: htmlBody, 
    cc: 'YourAdminEmail@gmail.com',
    name: "Automated invoice"
  });

}

function customRound(number) {
  // Check if the number is negative
  var isNegative = number < 0;
  
  // Take the absolute value for easier manipulation
  number = Math.abs(number);
  
  // Get the decimal part of the number
  var decimalPart = number - Math.floor(number);
  
  // If decimal part is less than 0.5, round down; otherwise, round up
  var roundedNumber = decimalPart < 0.5 ? Math.floor(number) : Math.ceil(number);
  
  // If the original number was negative, make the rounded number negative
  if (isNegative) {
    roundedNumber = -roundedNumber;
  }
  
  return roundedNumber;
}

// Convert Total Price in words.
function NumInWords(num) {
    var ones = ["", "One ", "Two ", "Three ", "Four ", "Five ", "Six ", "Seven ", "Eight ", "Nine ", "Ten ", "Eleven ", "Twelve ", "Thirteen ", "Fourteen ", "Fifteen ", "Sixteen ", "Seventeen ", "Eighteen ", "Nineteen "];
    var tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
    if ((num = num.toString()).length > 9) return "Overflow: Maximum 9 digits supported";
    n = ("000000000" + num).substr(-9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
    if (!n) return;
    var str = "";
    str += n[1] != 0 ? (ones[Number(n[1])] || tens[n[1][0]] + " " + ones[n[1][1]]) + "Crore " : "";
    str += n[2] != 0 ? (ones[Number(n[2])] || tens[n[2][0]] + " " + ones[n[2][1]]) + "Lakh " : "";
    str += n[3] != 0 ? (ones[Number(n[3])] || tens[n[3][0]] + " " + ones[n[3][1]]) + "Thousand " : "";
    str += n[4] != 0 ? (ones[Number(n[4])] || tens[n[4][0]] + " " + ones[n[4][1]]) + "Hundred " : "";
    str += n[5] != 0 ? (str != "" ? "And " : "") + (ones[Number(n[5])] || tens[n[5][0]] + " " + ones[n[5][1]]) : "";
    return str;
}

// Create Year-Month wise folder in the drive if not exists.
function getYearMonthFolder() {
  var mainFolderName = 'Invoices';
  // Get current year and month
  var today = new Date();
  var year = today.getFullYear();
  var month = today.getMonth() + 1; // Month is zero-based
  
  // Define folder names
  var yearFolderName = year;
  var monthFolderNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  
  var mainFolder = DriveApp.getFoldersByName(mainFolderName);
  if (!mainFolder.hasNext()) {
    mainFolder = DriveApp.createFolder(mainFolderName);
  } else {
    mainFolder = mainFolder.next();
  }
  // Check if year folder exists, if not, create it
  var yearFolder = mainFolder.getFoldersByName(yearFolderName);
  var existingYearFolder;
  if (!yearFolder.hasNext()) {
    var newYearFolder = mainFolder.createFolder(yearFolderName);
    existingYearFolder = newYearFolder;
  } else {
    existingYearFolder = yearFolder.next();
  }
  
  // Check if month folder exists, if not, create it
  var monthFolder = existingYearFolder.getFoldersByName(monthFolderNames[month - 1]);
  var existingMonthFolder;
  if (!monthFolder.hasNext()) {
    var newMonthFolder = existingYearFolder.createFolder(monthFolderNames[month - 1]);
    existingMonthFolder = newMonthFolder;
  } else {
    existingMonthFolder = monthFolder.next();
  }

  return {
    monthFolder: existingMonthFolder, // Return the modified data object
    yearFolder: existingYearFolder // Return the grand total
  };
}

