// Configuration
const DB_ID = '1TlcVa6NAX-2o4qGzfb1jsw3tD-R8L8K5jucwLcD_mGA';
const TARGET_FOLDER_ID = '1PmESdk0hdF0H38eTOOJ-uOlhAOMz6-by';

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Travel Orders')
      .addItem('Open Dashboard', 'showDashboard')
      .addItem('New Travel Order', 'showNewToForm')
      .addToUi();
  } catch (error) {
    Logger.log('Error in onOpen: ' + error.message);
  }
}

/**
 * Web app entry point
 */
function doGet() {
  try {
    return showDashboard();
  } catch (error) {
    Logger.log('Error in doGet: ' + error.message);
    return HtmlService.createHtmlOutput(
      '<html><body><h2>Error Loading Application</h2>' +
      '<p>Error: ' + error.message + '</p>' +
      '<p>Please ensure the spreadsheet is properly configured and you have access.</p></body></html>'
    ).setTitle('Error');
  }
}

/**
 * Generates unique Travel Order ID
 * Format: RO2.2-YYYY-JO#####
 */
function generateToId(datePrepared) {
  try {
    const sheet = SpreadsheetApp.openById(DB_ID).getSheetByName('Travel Orders');
    if (!sheet) {
      throw new Error('Sheet "Travel Orders" not found');
    }

    const year = Utilities.formatDate(new Date(datePrepared), 'Asia/Manila', 'yyyy');
    const data = sheet.getDataRange().getValues();

    let maxSerial = 0;
    for (let i = 1; i < data.length; i++) {
      const toId = data[i][0];
      if (toId && toId.toString().startsWith('RO2.2-' + year)) {
        const serialMatch = toId.toString().match(/-JO(\d+)$/);
        if (serialMatch) {
          const serial = parseInt(serialMatch[1]);
          if (serial > maxSerial) maxSerial = serial;
        }
      }
    }

    const newSerial = (maxSerial + 1).toString().padStart(5, '0');
    return `RO2.2-${year}-JO${newSerial}`;
  } catch (error) {
    Logger.log('Error in generateToId: ' + error.message);
    throw new Error('Failed to generate Travel Order ID: ' + error.message);
  }
}

/**
 * Saves new Travel Order with validation
 */
function saveTravelOrder(form) {
  try {
    const datePrepared = new Date(form.datePrepared);
    const start = new Date(form.inclusiveStart);
    const end = new Date(form.inclusiveEnd);
    const submission = new Date(form.dateSubmission);
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Validate Date Prepared
    if (datePrepared > today) {
      throw new Error('Date Prepared cannot be in the future.');
    }

    const sheet = SpreadsheetApp.openById(DB_ID).getSheetByName('Travel Orders');
    if (!sheet) {
      throw new Error('Sheet "Travel Orders" not found');
    }

    const data = sheet.getDataRange().getValues();
    let latestDatePrepared = null;

    // Find latest Date Prepared
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) {
        const existingDate = new Date(data[i][1]);
        if (!latestDatePrepared || existingDate > latestDatePrepared) {
          latestDatePrepared = existingDate;
        }
      }
    }

    if (latestDatePrepared && datePrepared < latestDatePrepared) {
      throw new Error('Date Prepared must be greater than or equal to the latest Date Prepared in the database.');
    }

    // Validate dates
    if (start < datePrepared) {
      throw new Error('Start Date cannot be before Date Prepared.');
    }

    if (end < start) {
      throw new Error('End Date cannot be before Start Date.');
    }

    if (submission < end) {
      throw new Error('Submission date cannot be before End Date.');
    }

    const toId = generateToId(datePrepared);
    const userEmail = Session.getActiveUser().getEmail();
    const now = new Date();

    // Append new row
    sheet.appendRow([
      toId,
      Utilities.formatDate(datePrepared, 'Asia/Manila', 'yyyy-MM-dd'),
      Utilities.formatDate(start, 'Asia/Manila', 'yyyy-MM-dd'),
      Utilities.formatDate(end, 'Asia/Manila', 'yyyy-MM-dd'),
      form.destination,
      form.purpose,
      form.requestedBy,
      form.requestingOfficer,
      Utilities.formatDate(submission, 'Asia/Manila', 'yyyy-MM-dd'),
      'Pending',
      '',
      userEmail,
      now,
      userEmail,
      now
    ]);

    return { success: true, toId: toId };
  } catch (error) {
    Logger.log('Error in saveTravelOrder: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Updates existing Travel Order
 */
function updateTravelOrder(form) {
  try {
    const start = new Date(form.inclusiveStart);
    const end = new Date(form.inclusiveEnd);
    const submission = new Date(form.dateSubmission);
    const datePrepared = new Date(form.datePrepared);

    // Validate dates
    if (start < datePrepared) {
      throw new Error('Start Date cannot be before Date Prepared.');
    }

    if (end < start) {
      throw new Error('End Date cannot be before Start Date.');
    }

    if (submission < end) {
      throw new Error('Submission date cannot be before End Date.');
    }

    const sheet = SpreadsheetApp.openById(DB_ID).getSheetByName('Travel Orders');
    if (!sheet) {
      throw new Error('Sheet "Travel Orders" not found');
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === form.toId) {
        const userEmail = Session.getActiveUser().getEmail();
        const now = new Date();

        sheet.getRange(i + 1, 3).setValue(Utilities.formatDate(start, 'Asia/Manila', 'yyyy-MM-dd'));
        sheet.getRange(i + 1, 4).setValue(Utilities.formatDate(end, 'Asia/Manila', 'yyyy-MM-dd'));
        sheet.getRange(i + 1, 5).setValue(form.destination);
        sheet.getRange(i + 1, 6).setValue(form.purpose);
        sheet.getRange(i + 1, 7).setValue(form.requestedBy);
        sheet.getRange(i + 1, 8).setValue(form.requestingOfficer);
        sheet.getRange(i + 1, 9).setValue(Utilities.formatDate(submission, 'Asia/Manila', 'yyyy-MM-dd'));
        sheet.getRange(i + 1, 14).setValue(userEmail);
        sheet.getRange(i + 1, 15).setValue(now);

        return { success: true };
      }
    }

    throw new Error('Travel Order not found.');
  } catch (error) {
    Logger.log('Error in updateTravelOrder: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Retrieves all Travel Orders
 */
function getAllTravelOrders() {
  try {
    const sheet = SpreadsheetApp.openById(DB_ID).getSheetByName('Travel Orders');
    if (!sheet) {
      Logger.log('Sheet "Travel Orders" not found');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const orders = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        orders.push({
          toId: data[i][0],
          datePrepared: data[i][1],
          inclusiveStart: data[i][2],
          inclusiveEnd: data[i][3],
          destination: data[i][4],
          purpose: data[i][5],
          requestedBy: data[i][6],
          requestingOfficer: data[i][7],
          dateSubmission: data[i][8],
          status: data[i][9],
          pdfLink: data[i][10]
        });
      }
    }

    return orders;
  } catch (error) {
    Logger.log('Error in getAllTravelOrders: ' + error.message);
    return [];
  }
}

/**
 * Retrieves a specific Travel Order by ID
 */
function getTravelOrderById(toId) {
  try {
    const sheet = SpreadsheetApp.openById(DB_ID).getSheetByName('Travel Orders');
    if (!sheet) {
      Logger.log('Sheet "Travel Orders" not found');
      return null;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === toId) {
        return {
          toId: data[i][0],
          datePrepared: data[i][1],
          inclusiveStart: data[i][2],
          inclusiveEnd: data[i][3],
          destination: data[i][4],
          purpose: data[i][5],
          requestedBy: data[i][6],
          requestingOfficer: data[i][7],
          dateSubmission: data[i][8],
          status: data[i][9],
          pdfLink: data[i][10]
        };
      }
    }

    return null;
  } catch (error) {
    Logger.log('Error in getTravelOrderById: ' + error.message);
    return null;
  }
}

/**
 * Gets the latest Date Prepared from database
 */
function getLatestDatePrepared() {
  try {
    const sheet = SpreadsheetApp.openById(DB_ID).getSheetByName('Travel Orders');
    if (!sheet) {
      Logger.log('Sheet "Travel Orders" not found');
      return null;
    }

    const data = sheet.getDataRange().getValues();
    let latestDate = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) {
        const existingDate = new Date(data[i][1]);
        if (!latestDate || existingDate > latestDate) {
          latestDate = existingDate;
        }
      }
    }

    return latestDate ? Utilities.formatDate(latestDate, 'Asia/Manila', 'yyyy-MM-dd') : null;
  } catch (error) {
    Logger.log('Error in getLatestDatePrepared: ' + error.message);
    return null;
  }
}
