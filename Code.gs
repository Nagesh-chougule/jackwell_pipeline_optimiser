/**
 * Google Apps Script Web App - Pipeline Path Optimiser Integration
 * 
 * INSTRUCTIONS FOR DEPLOYMENT:
 * 1. Open Google Sheets (https://sheets.google.com).
 * 2. Create a new empty spreadsheet or open an existing one.
 * 3. In the top menu, go to Extensions > Apps Script.
 * 4. Delete any code in the editor and paste this entire code.
 * 5. Click the "Save" icon (or press Ctrl+S).
 * 6. In the top right, click "Deploy" > "New deployment".
 * 7. Click the gear icon next to "Select type" and select "Web app".
 * 8. Enter a description (e.g., "Pipeline Optimiser API").
 * 9. Set "Execute as" to "Me (your email)".
 * 10. Set "Who has access" to "Anyone" (this is required so the HTML page can send requests without needing authentication).
 * 11. Click "Deploy".
 * 12. If prompted, click "Authorize access" and choose your Google account. Click "Advanced" > "Go to ... (unsafe)" and then "Allow" to grant permissions.
 * 13. Copy the "Web app URL" (it will look like: https://script.google.com/macros/s/.../exec).
 * 14. Paste this URL into the "Google Apps Script Web App URL" input field in the Pipeline Optimiser UI.
 */

function doPost(e) {
  try {
    // Parse the incoming JSON payload.
    // We send payload as a plain text string to avoid CORS preflight (OPTIONS) pre-requisites.
    var payload;
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      throw new Error("No payload data received.");
    }
    
    // Get the active spreadsheet sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Define the columns mapping in the spreadsheet
    var headers = [
      "Timestamp",
      "Jackwell Name",
      "Jackwell Latitude",
      "Jackwell Longitude",
      "Jackwell Elevation (m)",
      "WTP Name",
      "WTP Latitude",
      "WTP Longitude",
      "WTP Elevation (m)",
      "Design Flow Q (m³/hr)",
      "Pipe Material",
      "Target Velocity (m/s)",
      "Terrain Factor",
      "Route Strategy",
      "Minor Loss Coeff. k",
      "Straight-line Distance (km)",
      "Route Length (km)",
      "Pipeline Length (km)",
      "Pipe DN (mm)",
      "Actual Velocity (m/s)",
      "Static Head (m)",
      "Friction Loss (m)",
      "Minor Loss (m)",
      "Total Dynamic Head TDH (m)",
      "Motor Power (kW)",
      "Estimated Cost (₹ Cr)"
    ];
    
    // If the sheet is brand new/empty, write and format the headers first
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#2c2c2a");
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
      // Auto-resize columns for readability
      for (var col = 1; col <= headers.length; col++) {
        sheet.autoResizeColumn(col);
      }
    }
    
    // Format the date/time to local timezone format
    var timestamp = payload.timestamp ? new Date(payload.timestamp) : new Date();
    
    // Map fields from JSON structure to array elements
    var rowData = [
      timestamp,
      payload.jname || "",
      payload.jlat || 0,
      payload.jlon || 0,
      payload.jel || 0,
      payload.wname || "",
      payload.wlat || 0,
      payload.wlon || 0,
      payload.wel || 0,
      payload.flow_q || 0,
      payload.material || "",
      payload.target_velocity || 0,
      payload.terrain_factor || 0,
      payload.route_strategy || "",
      payload.minor_loss_k || 0,
      payload.straight_dist_km || 0,
      payload.route_length_km || 0,
      payload.pipe_length_km || 0,
      payload.pipe_dn_mm || 0,
      payload.actual_velocity || 0,
      payload.static_head_m || 0,
      payload.friction_loss_m || 0,
      payload.minor_loss_m || 0,
      payload.tdh_m || 0,
      payload.motor_kw || 0,
      payload.total_cost_cr || 0
    ];
    
    // Append the row
    sheet.appendRow(rowData);
    
    // Return standard success response
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Row appended successfully!"
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      "Access-Control-Allow-Origin": "*"
    });
    
  } catch (error) {
    // Return error response
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      "Access-Control-Allow-Origin": "*"
    });
  }
}
