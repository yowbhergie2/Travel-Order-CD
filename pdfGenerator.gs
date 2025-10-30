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
          body { font-family: Arial, sans-serif; padding: 40px; }
          .header { text-align: center; margin-bottom: 30px; }
          .header h1 { margin: 5px 0; font-size: 24px; }
          .header h2 { margin: 5px 0; font-size: 18px; font-weight: normal; }
          .content { margin: 20px 0; }
          .field { margin: 15px 0; }
          .label { font-weight: bold; display: inline-block; width: 200px; }
          .value { display: inline-block; }
          .footer { margin-top: 50px; }
          .signature { margin-top: 30px; }
          .signature-line { border-top: 1px solid #000; width: 250px; display: inline-block; }
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
            <p>Approved by:</p>
            <br><br>
            <div class="signature-line"></div>
            <p>${order.requestingOfficer}</p>
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
    return { success: false, error: error.message };
  }
}

function formatDisplayDate(dateString) {
  const date = new Date(dateString);
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
  return `${months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
}

function formatInclusiveDates(startString, endString) {
  const start = new Date(startString);
  const end = new Date(endString);
  
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
  
  if (start.getTime() === end.getTime()) {
    return `${months[start.getMonth()]} ${start.getDate()}, ${start.getFullYear()}`;
  }
  
  if (start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear()) {
    return `${months[start.getMonth()]} ${start.getDate()}–${end.getDate()}, ${start.getFullYear()}`;
  }
  
  return `${months[start.getMonth()]} ${start.getDate()}, ${start.getFullYear()} – ${months[end.getMonth()]} ${end.getDate()}, ${end.getFullYear()}`;
}
