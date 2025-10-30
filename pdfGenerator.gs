/**
 * Generates PDF for a Travel Order and saves it to Drive
 */
function generateTravelOrderPdf(toId) {
  try {
    const order = getTravelOrderById(toId);
    if (!order) {
      throw new Error('Travel Order not found.');
    }

    const inclusiveDatesFormatted = formatInclusiveDates(order.inclusiveStart, order.inclusiveEnd);
    const datePreparedFormatted = formatDisplayDate(order.datePrepared);
    const dateSubmissionFormatted = formatDisplayDate(order.dateSubmission);

    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 40px;
            line-height: 1.6;
          }
          .header {
            text-align: center;
            margin-bottom: 30px;
            border-bottom: 2px solid #333;
            padding-bottom: 20px;
          }
          .header h1 {
            margin: 5px 0;
            font-size: 28px;
            color: #2c3e50;
          }
          .header h2 {
            margin: 5px 0;
            font-size: 20px;
            font-weight: normal;
            color: #34495e;
          }
          .content {
            margin: 20px 0;
          }
          .field {
            margin: 15px 0;
            padding: 10px;
            background: #f8f9fa;
            border-left: 4px solid #3498db;
          }
          .label {
            font-weight: bold;
            display: inline-block;
            width: 250px;
            color: #2c3e50;
          }
          .value {
            display: inline-block;
            color: #34495e;
          }
          .footer {
            margin-top: 60px;
            padding-top: 20px;
            border-top: 1px solid #ccc;
          }
          .signature {
            margin-top: 40px;
          }
          .signature-line {
            border-top: 2px solid #000;
            width: 300px;
            display: inline-block;
            margin-top: 40px;
          }
          .signature p {
            margin: 5px 0;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>TRAVEL ORDER</h1>
          <h2>${toId}</h2>
        </div>

        <div class="content">
          <div class="field">
            <span class="label">Date Prepared:</span>
            <span class="value">${datePreparedFormatted}</span>
          </div>

          <div class="field">
            <span class="label">Inclusive Dates:</span>
            <span class="value">${inclusiveDatesFormatted}</span>
          </div>

          <div class="field">
            <span class="label">Destination:</span>
            <span class="value">${order.destination}</span>
          </div>

          <div class="field">
            <span class="label">Purpose:</span>
            <span class="value">${order.purpose}</span>
          </div>

          <div class="field">
            <span class="label">Requested By:</span>
            <span class="value">${order.requestedBy}</span>
          </div>

          <div class="field">
            <span class="label">Requesting Officer:</span>
            <span class="value">${order.requestingOfficer}</span>
          </div>

          <div class="field">
            <span class="label">Date of Submission of Travel Report:</span>
            <span class="value">${dateSubmissionFormatted}</span>
          </div>
        </div>

        <div class="footer">
          <div class="signature">
            <p><strong>Approved by:</strong></p>
            <div class="signature-line"></div>
            <p><strong>${order.requestingOfficer}</strong></p>
            <p>Requesting Officer</p>
          </div>
        </div>
      </body>
      </html>
    `;

    const blob = Utilities.newBlob(htmlContent, 'text/html', `${toId}.html`);
    const pdfBlob = blob.getAs('application/pdf');

    const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    const file = folder.createFile(pdfBlob);
    file.setName(`${toId}.pdf`);

    const pdfUrl = file.getUrl();

    // Update the sheet with PDF link and status
    const sheet = SpreadsheetApp.openById(DB_ID).getSheetByName('Travel Orders');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === toId) {
        sheet.getRange(i + 1, 10).setValue('Completed');
        sheet.getRange(i + 1, 11).setValue(pdfUrl);
        break;
      }
    }

    return { success: true, pdfUrl: pdfUrl };
  } catch (error) {
    Logger.log('Error in generateTravelOrderPdf: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Formats a date string into human-readable format
 * Example: "2025-10-05" -> "October 5, 2025"
 */
function formatDisplayDate(dateString) {
  const date = new Date(dateString);
  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];
  return `${months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
}

/**
 * Formats inclusive dates range
 * Examples:
 * - "October 5, 2025" (single day)
 * - "October 5–9, 2025" (same month)
 * - "October 30, 2025 – November 2, 2025" (cross-month/year)
 */
function formatInclusiveDates(startString, endString) {
  const start = new Date(startString);
  const end = new Date(endString);

  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];

  // Same day
  if (start.getTime() === end.getTime()) {
    return `${months[start.getMonth()]} ${start.getDate()}, ${start.getFullYear()}`;
  }

  // Same month and year
  if (start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear()) {
    return `${months[start.getMonth()]} ${start.getDate()}–${end.getDate()}, ${start.getFullYear()}`;
  }

  // Different month or year
  return `${months[start.getMonth()]} ${start.getDate()}, ${start.getFullYear()} – ${months[end.getMonth()]} ${end.getDate()}, ${end.getFullYear()}`;
}
