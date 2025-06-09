// --- Global Configuration ---
const SUMMARY_SHEET_NAME = "סיכום יומי"; // Name of the summary sheet
const DATE_INPUT_CELL = "B2"; // Cell where the target date for daily report is entered
const OUTPUT_HEADER_ROW = 5; // Starting row for output headers in the summary sheet
const OUTPUT_START_ROW = 6; // Starting row for data in the summary sheet

// --- Client Sheet Configuration ---
const CLIENT_SHEET_HEADER_ROWS = 1; // Number of header rows on client-specific sheets

// Headers for Client Sheets (e.g., "מאוחדת"). This defines the order of data written.
const CLIENT_SHEET_HEADERS = [
  "נהג", // Driver's email
  "חותמת זמן", // Timestamp of submission
  "דרייב / פיזי", // Drive / Physical delivery type
  "מספר שורה", // Line number (kept blank)
  "מספר התחייבות", // Commitment number
  "מספר זיהוי", // Identification number
  "שם פרטי", // First name
  "שם משפחה", // Last name
  "תאור הטיפול", // Description of treatment/delivery
  "תאריך", // Delivery date
  "כמות", // Quantity
  "מחיר", // Price (kept blank)
  "סכום" // Sum (kept blank)
];

const DATE_COLUMN_INDEX_ON_CLIENT_SHEET = CLIENT_SHEET_HEADERS.indexOf("תאריך"); // Index of the 'תאריך' column (0-indexed)

const OUTPUT_COLUMN_HEADERS_FOR_SUMMARY = [
  "שם לקוח", // Client name (added as the first column in summary output)
  ...CLIENT_SHEET_HEADERS // All client sheet headers follow
];

// --- Form Responses Sheet Configuration ---
const FORM_RESPONSES_SHEET_NAME = "הזנות"; // Name of the sheet linked to Google Forms responses

// --- End Global Configuration ---

/**
 * Creates a custom menu in the spreadsheet when it's opened, allowing manual script execution.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('כלי משלוחים')
      .addItem('הפק דוח משלוחים יומי', 'generateDailyDeliveryReport')
      .addToUi();
  Logger.log("Custom menu 'כלי משלוחים' created/updated.");
}

/**
 * Automatically triggered on Google Forms submission.
 * Reads new form data, appends it to the relevant client sheet and the main summary sheet.
 * @param {Object} e Event object containing form submission details.
 */
function onFormSubmitTrigger(e) {
  Logger.log("--- onFormSubmitTrigger started ---");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    if (!e || !e.range || !e.range.getSheet()) {
      Logger.log(`Error: Invalid trigger event object. e: ${JSON.stringify(e)}`);
      return;
    }

    const formResponsesSheet = e.range.getSheet();
    if (formResponsesSheet.getName() !== FORM_RESPONSES_SHEET_NAME) {
      Logger.log(`Trigger activated from non-form responses sheet: "${formResponsesSheet.getName()}". Skipping.`);
      return;
    }

    const newRowData = e.range.getValues()[0];
    Logger.log(`Raw data from form responses sheet (row ${e.range.getRow()}): ${JSON.stringify(newRowData)}`);

    // Map form response columns to variables based on their index in the raw data
    const timestampRaw = newRowData[0]; // Column A: Timestamp
    const email = String(newRowData[1] || '').trim(); // Column B: Driver's email
    const clientName = String(newRowData[2] || '').trim(); // Column C: Client name
    const drivePhysical = String(newRowData[3] || '').trim(); // Column D: Drive / Physical
    const lineNumber = ''; // Column E: Line number (intentionally left blank)
    const commitmentNumber = String(newRowData[4] || '').trim(); // Column F: Commitment Number
    const identificationNumber = String(newRowData[5] || '').trim(); // Column G: Identification Number
    const firstName = String(newRowData[6] || '').trim(); // Column H: First Name
    const lastName = String(newRowData[7] || '').trim(); // Column I: Last Name
    const fromLocationRaw = String(newRowData[8] || '').trim(); // Column J: "From where"
    const toLocationRaw = String(newRowData[9] || '').trim(); // Column K: "To where"
    const rawDateFromForm = newRowData[10]; // Column L: Date from form
    const rawQuantityString = String(newRowData[11] || '').trim(); // Column M: Quantity

    // Fields not present in the form, to be kept blank or derived
    const price = '';
    const sum = '';
    const description = `מ${fromLocationRaw} ל${toLocationRaw}`;

    // Process timestamp
    let formattedTimestamp = '';
    if (timestampRaw && timestampRaw instanceof Date) {
      formattedTimestamp = Utilities.formatDate(timestampRaw, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
    } else if (typeof timestampRaw === 'string' && timestampRaw !== '') {
      try {
          const parsedTimestamp = new Date(timestampRaw);
          if (!isNaN(parsedTimestamp.getTime())) {
             formattedTimestamp = Utilities.formatDate(parsedTimestamp, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
      } catch (e) {
          Logger.log(`Error parsing timestamp for "${timestampRaw}": ${e.message}`);
      }
    }

    // Process date for use (prioritizes form date)
    let dateToUse = null;
    let formattedDate = '';
    const spreadsheetTimeZone = ss.getSpreadsheetTimeZone();
    const hebrewDateFormat = "dd/MM/yyyy"; // Desired date format: 08/06/2025

    // Attempt to parse the date from the form
    if (rawDateFromForm && rawDateFromForm instanceof Date && !isNaN(rawDateFromForm.getTime())) {
        dateToUse = rawDateFromForm;
        Logger.log(`Using date from form (Date object): ${rawDateFromForm}`);
    } else if (typeof rawDateFromForm === 'string' && rawDateFromForm.trim() !== '') {
        try {
            const parsedFormDate = new Date(rawDateFromForm);
            if (!isNaN(parsedFormDate.getTime())) {
                dateToUse = parsedFormDate;
                Logger.log(`Using parsed string date from form: ${rawDateFromForm}`);
            }
        } catch (e) {
            Logger.log(`Error parsing string date from form "${rawDateFromForm}": ${e.message}`);
        }
    }

    // If form date is empty or invalid, use timestamp date
    if (!dateToUse && timestampRaw && timestampRaw instanceof Date && !isNaN(timestampRaw.getTime())) {
        dateToUse = timestampRaw;
        Logger.log(`Form date missing/invalid. Using timestamp date (Date object): ${timestampRaw}`);
    } else if (!dateToUse && typeof timestampRaw === 'string' && timestampRaw.trim() !== '') {
         try {
            const parsedTimestampDate = new Date(timestampRaw);
            if (!isNaN(parsedTimestampDate.getTime())) {
                dateToUse = parsedTimestampDate;
                Logger.log(`Form date missing/invalid. Using parsed timestamp string date: ${timestampRaw}`);
            }
        } catch (e) {
            Logger.log(`Error parsing timestamp string for date: ${e.message}`);
        }
    }

    // Format the date for spreadsheet use
    if (dateToUse) {
      formattedDate = Utilities.formatDate(dateToUse, spreadsheetTimeZone, hebrewDateFormat);
      Logger.log(`Final date used: ${formattedDate} (Source: ${dateToUse})`);
    } else {
      formattedDate = '';
      Logger.log(`No valid date found for the record.`);
    }

    // Process quantity (הלוך -> 1, הלוך-חזור -> 2)
    let quantity;
    if (rawQuantityString === "הלוך") {
        quantity = 1;
    } else if (rawQuantityString === "הלוך-חזור") {
        quantity = 2;
    } else {
        quantity = ''; // Leave blank if not "הלוך" or "הלוך-חזור"
    }

    Logger.log(`Processed data: Client=${clientName}, Driver (email)=${email}, Timestamp=${formattedTimestamp}, Drive/Physical=${drivePhysical}, Line Number=${lineNumber}, Commitment Number=${commitmentNumber}, Identification Number=${identificationNumber}, First Name=${firstName}, Last Name=${lastName}, Description=${description}, Date=${formattedDate}, Quantity=${quantity}, Price=${price}, Sum=${sum}`);

    // Validate essential fields
    if (!clientName || clientName === '' ||
        !firstName || firstName === '' ||
        !lastName || lastName === '' ||
        !description || description === '') {
      Logger.log(`Warning: Mandatory fields (Client, First Name, Last Name, Description) are missing/partial. Record will be added with incomplete data.`);
    }

    // Find the corresponding client sheet
    const targetClientSheet = ss.getSheetByName(clientName);
    if (!targetClientSheet) {
      Logger.log(`Error: Client sheet "${clientName}" (from form) not found in the main spreadsheet. Record not added.`);
      return;
    }
    Logger.log(`Client sheet "${clientName}" found successfully.`);

    // Construct the row for the client sheet based on CLIENT_SHEET_HEADERS order
    const clientSheetRow = [
      email,
      formattedTimestamp,
      drivePhysical,
      lineNumber,
      commitmentNumber,
      identificationNumber,
      firstName,
      lastName,
      description,
      formattedDate,
      quantity,
      price,
      sum
    ];

    // Write data to the client sheet
    const targetRowClientSheet = targetClientSheet.getLastRow() + 1;
    const finalTargetRowClientSheet = Math.max(targetRowClientSheet, CLIENT_SHEET_HEADER_ROWS + 1);

    targetClientSheet.getRange(finalTargetRowClientSheet, 1, 1, clientSheetRow.length).setValues([clientSheetRow]);
    Logger.log(`New row appended to client sheet "${clientName}" at row ${finalTargetRowClientSheet}.`);

    // Append the row to the main summary sheet
    const summarySheetOutputRow = [
      clientName,
      email,
      formattedTimestamp,
      drivePhysical,
      lineNumber,
      commitmentNumber,
      identificationNumber,
      firstName,
      lastName,
      description,
      formattedDate,
      quantity,
      price,
      sum
    ];

    const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
    if (summarySheet) {
      const targetRowSummarySheet = summarySheet.getLastRow() + 1;
      const finalTargetRowSummarySheet = Math.max(targetRowSummarySheet, OUTPUT_START_ROW);

      summarySheet.getRange(finalTargetRowSummarySheet, 1, 1, summarySheetOutputRow.length).setValues([summarySheetOutputRow]);
      Logger.log(`New data successfully appended to summary sheet "${SUMMARY_SHEET_NAME}" starting at row ${finalTargetRowSummarySheet}.`);

      for (let i = 1; i <= OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.length; i++) {
        summarySheet.autoResizeColumn(i);
      }
      Logger.log("Summary sheet columns auto-resized.");

      // Format date column in summary sheet
      const outputDateColumnIndex = OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.indexOf("תאריך") + 1;
      if (outputDateColumnIndex > 0) {
        summarySheet.getRange(finalTargetRowSummarySheet, outputDateColumnIndex, 1, 1).setNumberFormat(hebrewDateFormat);
        Logger.log(`Date column (output column ${outputDateColumnIndex}) formatted to "${hebrewDateFormat}".`);
      }

      // Format timestamp column in summary sheet
      const outputTimestampColumnIndex = OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.indexOf("חותמת זמן") + 1;
      if (outputTimestampColumnIndex > 0) {
        summarySheet.getRange(finalTargetRowSummarySheet, outputTimestampColumnIndex, 1, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
        Logger.log(`Timestamp column (output column ${outputTimestampColumnIndex}) formatted to "yyyy-MM-dd HH:mm:ss".`);
      }

    } else {
      Logger.log(`Error: Summary sheet "${SUMMARY_SHEET_NAME}" not found, cannot append data.`);
    }

    Logger.log(`Data successfully saved for ${clientName}.`);

  } catch (e) {
    Logger.log(`--- Script Error in onFormSubmitTrigger --- \nError Name: ${e.name}\nError Message: ${e.message}\nStack Trace: ${e.stack}\n-------------------------`);
  } finally {
    Logger.log("--- onFormSubmitTrigger finished ---");
  }
}

/**
 * Main function to generate the daily delivery report.
 * This function reads data from all client sheets, filters deliveries by the selected date,
 * and writes the results to the SUMMARY_SHEET_NAME.
 */
function generateDailyDeliveryReport() {
  Logger.log("--- generateDailyDeliveryReport started ---");
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`Active Spreadsheet: ID = ${ss.getId()}, Name = "${ss.getName()}"`);

    const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
    if (!summarySheet) {
      const errorMessage = `Sheet named "${SUMMARY_SHEET_NAME}" not found. Please create it or check SUMMARY_SHEET_NAME in the script.`;
      Logger.log(`Error: ${errorMessage}`);
      ui.alert("Configuration Error", errorMessage, ui.ButtonSet.OK);
      return;
    }
    Logger.log(`Summary sheet "${SUMMARY_SHEET_NAME}" found and opened successfully.`);

    // Critical check before clearing content to prevent data loss
    if (summarySheet.getName() !== SUMMARY_SHEET_NAME) {
      const errorMessage = `Critical Warning: Script attempted to clear sheet "${summarySheet.getName()}" which is not the configured summary sheet "${SUMMARY_SHEET_NAME}". Clear operation aborted to prevent data loss.`;
      Logger.log(errorMessage);
      ui.alert("Critical Error", errorMessage, ui.ButtonSet.OK);
      return;
    }

    const targetDateValue = summarySheet.getRange(DATE_INPUT_CELL).getValue();
    Logger.log(`Raw date value read from cell ${DATE_INPUT_CELL} in "${SUMMARY_SHEET_NAME}": ${targetDateValue} (Type: ${typeof targetDateValue})`);

    if (!targetDateValue || !(targetDateValue instanceof Date)) {
      const errorMessage = `Please enter a valid date in cell ${DATE_INPUT_CELL} in sheet "${SUMMARY_SHEET_NAME}". The current value is not recognized as a date.`;
      Logger.log(`Input error: ${errorMessage}. Value was: ${targetDateValue}`);
      ui.alert("Input Error", errorMessage, ui.ButtonSet.OK);
      return;
    }

    // Normalize target date to midnight for consistent comparison
    const targetDate = new Date(targetDateValue.getFullYear(), targetDateValue.getMonth(), targetDateValue.getDate());
    Logger.log(`Normalized target date for comparison (midnight): ${Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")}`);

    const allSheets = ss.getSheets();
    Logger.log(`Total sheets found in the spreadsheet: ${allSheets.length}`);

    const compiledDeliveries = [];
    const numClientSheetColumnsToFetch = CLIENT_SHEET_HEADERS.length;

    Logger.log(`Configured to read ${numClientSheetColumnsToFetch} columns from each client sheet.`);

    for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      if (sheetName === SUMMARY_SHEET_NAME || sheetName === FORM_RESPONSES_SHEET_NAME) {
        Logger.log(`Skipping sheet: "${sheetName}" as it's a report destination or form responses sheet.`);
        continue;
      }

      Logger.log(`--- Processing client sheet: "${sheetName}" ---`);
      const maxRows = sheet.getMaxRows();

      if (maxRows <= CLIENT_SHEET_HEADER_ROWS) {
          Logger.log(`Sheet "${sheetName}" has no data rows available below headers (Header rows: ${CLIENT_SHEET_HEADER_ROWS}, Max rows: ${maxRows}). Skipping.`);
          continue;
      }

      const rowsToRead = sheet.getLastRow() - CLIENT_SHEET_HEADER_ROWS; // Read only up to the last row with content
      if (rowsToRead <= 0) {
          Logger.log(`No data rows to read in sheet "${sheetName}" after headers. Skipping.`);
          continue;
      }

      const dataRange = sheet.getRange(
        CLIENT_SHEET_HEADER_ROWS + 1,
        1,
        rowsToRead,
        numClientSheetColumnsToFetch
      );
      Logger.log(`Reading data from range: ${dataRange.getA1Notation()} in sheet "${sheetName}".`);
      const values = dataRange.getValues();
      Logger.log(`Found ${values.length} data rows (excluding headers) in sheet "${sheetName}".`);

      let deliveriesFoundInThisSheet = 0;
      for (const row of values) {
        // Ensure the row is not entirely empty and has enough columns for the date
        if (row.every(cell => cell === "") || row.length <= DATE_COLUMN_INDEX_ON_CLIENT_SHEET) {
          continue;
        }

        const deliveryDateValue = row[DATE_COLUMN_INDEX_ON_CLIENT_SHEET];

        if (deliveryDateValue && deliveryDateValue instanceof Date) {
          // Normalize delivery date to midnight for consistent comparison
          const deliveryDate = new Date(deliveryDateValue.getFullYear(), deliveryDateValue.getMonth(), deliveryDateValue.getDate());

          if (deliveryDate.getTime() === targetDate.getTime()) {
            const outputRow = [];
            outputRow.push(sheetName); // Client name is the first column in the summary

            for (let i = 0; i < CLIENT_SHEET_HEADERS.length; i++) {
                outputRow.push(row[i] !== undefined ? row[i] : "");
            }
            compiledDeliveries.push(outputRow);
            deliveriesFoundInThisSheet++;
          }
        } else if (deliveryDateValue) {
           // Log problematic date values for debugging
           Logger.log(`Skipping row in sheet "${sheetName}" due to invalid or non-date value in date column (index ${DATE_COLUMN_INDEX_ON_CLIENT_SHEET}). Value found: "${deliveryDateValue}". Raw row data: ${row.join("|")}`);
        }
      }
      Logger.log(`Processed ${values.length} rows in sheet "${sheetName}". Found ${deliveriesFoundInThisSheet} deliveries matching target date.`);
    }

    Logger.log(`Total compiled deliveries matching target date from all client sheets: ${compiledDeliveries.length}`);

    // --- Write results to the summary sheet ---
    // Clear previous content below headers
    const numRowsToClear = summarySheet.getMaxRows() - OUTPUT_HEADER_ROW + 1;
    if (numRowsToClear > 0) {
        summarySheet.getRange(OUTPUT_HEADER_ROW, 1, numRowsToClear, OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.length).clearContent();
        Logger.log(`Cleared previous content from summary sheet "${SUMMARY_SHEET_NAME}" starting at row ${OUTPUT_HEADER_ROW} for ${OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.length} columns.`);
    }

    // Write headers to summary sheet
    summarySheet.getRange(OUTPUT_HEADER_ROW, 1, 1, OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.length)
        .setValues([OUTPUT_COLUMN_HEADERS_FOR_SUMMARY])
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
    Logger.log(`Wrote new headers to row ${OUTPUT_HEADER_ROW} in "${SUMMARY_SHEET_NAME}". Headers: ${OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.join(", ")}`);

    if (compiledDeliveries.length > 0) {
      summarySheet.getRange(OUTPUT_START_ROW, 1, compiledDeliveries.length, compiledDeliveries[0].length).setValues(compiledDeliveries);
      Logger.log(`Wrote ${compiledDeliveries.length} delivery records to "${SUMMARY_SHEET_NAME}" starting at row ${OUTPUT_START_ROW}.`);

      for (let i = 1; i <= OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.length; i++) {
        summarySheet.autoResizeColumn(i);
      }
      Logger.log("Summary sheet columns auto-resized.");

      const hebrewDateFormat = "dd/MM/yyyy"; // Desired format for display in summary sheet
      const outputDateColumnIndex = OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.indexOf("תאריך") + 1;
      if (outputDateColumnIndex > 0) {
        summarySheet.getRange(OUTPUT_START_ROW, outputDateColumnIndex, compiledDeliveries.length, 1).setNumberFormat(hebrewDateFormat);
        Logger.log(`Date column (output column ${outputDateColumnIndex}) formatted to "${hebrewDateFormat}".`);
      }

      const outputTimestampColumnIndex = OUTPUT_COLUMN_HEADERS_FOR_SUMMARY.indexOf("חותמת זמן") + 1;
      if (outputTimestampColumnIndex > 0) {
        summarySheet.getRange(OUTPUT_START_ROW, outputTimestampColumnIndex, compiledDeliveries.length, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
        Logger.log(`Timestamp column (output column ${outputTimestampColumnIndex}) formatted to "yyyy-MM-dd HH:mm:ss".`);
      }

      ui.alert("Report Generated", `${compiledDeliveries.length} deliveries found for ${Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "dd/MM/yyyy")}.`, ui.ButtonSet.OK);
    } else {
      const noDataMessage = `No deliveries found for ${Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "dd/MM/yyyy")}.`;
      summarySheet.getRange(OUTPUT_START_ROW, 1).setValue(noDataMessage);
      Logger.log(noDataMessage);
      ui.alert("Report Generated", noDataMessage, ui.ButtonSet.OK);
    }

  } catch (e) {
    Logger.log(`--- Script Error --- \nError Name: ${e.name}\nError Message: ${e.message}\nStack Trace: ${e.stack}\n-------------------------`);
    ui.alert("Script Error", `An unexpected error occurred: ${e.message}\nMore details logged in script execution transcript.`, ui.ButtonSet.OK);
  } finally {
    Logger.log("--- generateDailyDailyReport finished ---");
  }
}
