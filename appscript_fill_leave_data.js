/**
 * Google Apps Script to fill jan_2026 sheet with expanded data from Ideal Order Quantity sheet
 *
 * CONFIGURATION SECTION - Modify these values as needed
 */

// Sheet names
const SOURCE_SHEET_NAME = "Ideal Order Quantity";
const DESTINATION_SHEET_NAME = "jan_2026";

// Source sheet settings
const SOURCE_START_ROW = 3;        // First data row (after headers)
const SOURCE_END_ROW = 224;        // Last data row
const SOURCE_HEADER_ROW = 2;       // Row containing size headers (XS, S, M, L, XL)

// Source column mappings (which columns to read)
const SOURCE_COL_B = 2;  // Column B
const SOURCE_COL_C = 3;  // Column C
const SOURCE_COL_D = 4;  // Column D
const SOURCE_COL_E = 5;  // Column E
const SOURCE_COL_F = 6;  // Column F
const SOURCE_COL_G = 7;  // Column G (XS count)
const SOURCE_COL_H = 8;  // Column H (S count)
const SOURCE_COL_I = 9;  // Column I (M count)
const SOURCE_COL_J = 10; // Column J (L count)
const SOURCE_COL_K = 11; // Column K (XL count)

// Destination sheet settings
const DEST_START_ROW = 2;          // Row to start writing data (modify as needed)

// Destination column mappings (where to write in jan_2026)
const DEST_COL_B = 2;  // Source B --> Dest B
const DEST_COL_C = 3;  // Source D --> Dest C
const DEST_COL_D = 4;  // Source E --> Dest D
const DEST_COL_E = 5;  // Source C --> Dest E
const DEST_COL_G = 7;  // Source F --> Dest G
const DEST_COL_H = 8;  // Size value --> Dest H

/**
 * Main function to fill the jan_2026 sheet
 */
function fillLeaveData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get sheets
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const destSheet = ss.getSheetByName(DESTINATION_SHEET_NAME);

  if (!sourceSheet) {
    throw new Error(`Source sheet "${SOURCE_SHEET_NAME}" not found`);
  }
  if (!destSheet) {
    throw new Error(`Destination sheet "${DESTINATION_SHEET_NAME}" not found`);
  }

  Logger.log("Starting data fill process...");

  // Read size headers from row 2 (columns G-K)
  const sizeHeaders = sourceSheet.getRange(SOURCE_HEADER_ROW, SOURCE_COL_G, 1, 5).getValues()[0];
  const sizes = {
    G: sizeHeaders[0], // XS
    H: sizeHeaders[1], // S
    I: sizeHeaders[2], // M
    J: sizeHeaders[3], // L
    K: sizeHeaders[4]  // XL
  };

  Logger.log(`Size headers: XS="${sizes.G}", S="${sizes.H}", M="${sizes.I}", L="${sizes.J}", XL="${sizes.K}"`);

  // Calculate number of rows to read
  const numRows = SOURCE_END_ROW - SOURCE_START_ROW + 1;

  // Read all source data at once for better performance
  const sourceData = sourceSheet.getRange(SOURCE_START_ROW, SOURCE_COL_B, numRows, SOURCE_COL_K - SOURCE_COL_B + 1).getValues();

  // Prepare output array
  const outputData = [];

  // Debugging counters
  let totalCellsProcessed = 0;
  let totalCellsWithValueGreaterThanZero = 0;
  let totalRowsGenerated = 0;
  let totalSumOfCellValues = 0;  // Track the sum to compare with SUMIF
  let decimalValues = [];  // Track any decimal values found
  let skippedCells = [];

  // Process each source row
  for (let i = 0; i < sourceData.length; i++) {
    const row = sourceData[i];
    const sourceRowNum = SOURCE_START_ROW + i;

    // Extract values from source columns
    const valueB = row[SOURCE_COL_B - SOURCE_COL_B]; // Column B
    const valueC = row[SOURCE_COL_C - SOURCE_COL_B]; // Column C
    const valueD = row[SOURCE_COL_D - SOURCE_COL_B]; // Column D
    const valueE = row[SOURCE_COL_E - SOURCE_COL_B]; // Column E
    const valueF = row[SOURCE_COL_F - SOURCE_COL_B]; // Column F

    // Extract counts from size columns (G-K)
    const counts = {
      G: row[SOURCE_COL_G - SOURCE_COL_B], // XS count
      H: row[SOURCE_COL_H - SOURCE_COL_B], // S count
      I: row[SOURCE_COL_I - SOURCE_COL_B], // M count
      J: row[SOURCE_COL_J - SOURCE_COL_B], // L count
      K: row[SOURCE_COL_K - SOURCE_COL_B]  // XL count
    };

    // Process each size column in order (G, H, I, J, K)
    const sizeColumns = ['G', 'H', 'I', 'J', 'K'];

    for (const col of sizeColumns) {
      const count = counts[col];
      const sizeValue = sizes[col];

      totalCellsProcessed++;

      // Debug: Check type and value
      const countNum = Number(count);

      // Skip if count is not a positive number
      if (count === "" || count === null || count === undefined || isNaN(countNum) || countNum <= 0) {
        if (count !== "" && count !== 0 && count !== null && count !== undefined) {
          // Log unexpected skipped values (first 20 only to avoid spam)
          if (skippedCells.length < 20) {
            skippedCells.push({row: sourceRowNum, col: col, value: count, type: typeof count});
          }
        }
        continue;
      }

      totalCellsWithValueGreaterThanZero++;
      totalSumOfCellValues += countNum;

      // Check if value is decimal (might cause rounding issues)
      if (countNum % 1 !== 0 && decimalValues.length < 20) {
        decimalValues.push({row: sourceRowNum, col: col, value: countNum});
      }

      // Create the specified number of rows
      for (let copy = 0; copy < countNum; copy++) {
        totalRowsGenerated++;
        // Create output row with mapped columns
        // Destination columns: B, C, D, E, G, H (skipping F)
        const outputRow = [];
        outputRow[DEST_COL_B - 1] = valueB;      // B --> B
        outputRow[DEST_COL_C - 1] = valueD;      // D --> C
        outputRow[DEST_COL_D - 1] = valueE;      // E --> D
        outputRow[DEST_COL_E - 1] = valueC;      // C --> E
        outputRow[DEST_COL_G - 1] = valueF;      // F --> G
        outputRow[DEST_COL_H - 1] = sizeValue;   // Size --> H

        // Fill empty columns with empty strings
        for (let j = 0; j < Math.max(DEST_COL_B, DEST_COL_C, DEST_COL_D, DEST_COL_E, DEST_COL_G, DEST_COL_H); j++) {
          if (outputRow[j] === undefined) {
            outputRow[j] = "";
          }
        }

        outputData.push(outputRow);
      }
    }

    // Log progress every 50 rows
    if ((i + 1) % 50 === 0) {
      Logger.log(`Processed ${i + 1} / ${numRows} source rows...`);
    }
  }

  // Log debugging statistics
  Logger.log("=== DEBUGGING STATISTICS ===");
  Logger.log(`Total cells processed (G-K across all rows): ${totalCellsProcessed}`);
  Logger.log(`Total cells with value > 0: ${totalCellsWithValueGreaterThanZero}`);
  Logger.log(`Total SUM of cell values > 0: ${totalSumOfCellValues} (compare with your SUMIF result: 766)`);
  Logger.log(`Total rows generated: ${totalRowsGenerated}`);
  Logger.log(`Total output rows in array: ${outputData.length}`);
  Logger.log(`Discrepancy: ${766 - totalSumOfCellValues} missing from sum, ${766 - totalRowsGenerated} missing from generated rows`);

  if (decimalValues.length > 0) {
    Logger.log("\nFirst 20 decimal values found (these get rounded down in for loops!):");
    decimalValues.forEach(cell => {
      Logger.log(`Row ${cell.row}, Col ${cell.col}: value=${cell.value}`);
    });
  }

  if (skippedCells.length > 0) {
    Logger.log("\nFirst 20 unexpected skipped cells:");
    skippedCells.forEach(cell => {
      Logger.log(`Row ${cell.row}, Col ${cell.col}: value="${cell.value}", type=${cell.type}`);
    });
  }

  // Write all data to destination sheet at once
  if (outputData.length > 0) {
    const numCols = outputData[0].length;
    destSheet.getRange(DEST_START_ROW, 1, outputData.length, numCols).setValues(outputData);
    Logger.log(`Successfully wrote ${outputData.length} rows to ${DESTINATION_SHEET_NAME} starting at row ${DEST_START_ROW}`);
  } else {
    Logger.log("No data to write (all counts were 0 or negative)");
  }

  const alertMessage = `Process complete!\n\nGenerated ${outputData.length} rows in "${DESTINATION_SHEET_NAME}" sheet.\n\nCells processed: ${totalCellsProcessed}\nCells > 0: ${totalCellsWithValueGreaterThanZero}\nSum of values: ${totalSumOfCellValues}\nExpected (SUMIF): 766\nDiscrepancy: ${766 - totalSumOfCellValues}\n\nCheck the Apps Script log (Ctrl/Cmd+Enter) for detailed breakdown.`;
  SpreadsheetApp.getUi().alert(alertMessage);
}
