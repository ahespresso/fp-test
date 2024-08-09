/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initialize();
  }
});

function initialize() {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("validateString").onclick = () => validate('string');
  document.getElementById("validateNumber").onclick = () => validate('number');
  document.getElementById("validateBoolean").onclick = () => validate('boolean');
  document.getElementById("validateDate").onclick = () => validate('date');
  document.getElementById("createDashboard").onclick = createDashboard;
  document.getElementById("trackKPIs").onclick = trackKPIs;
  document.getElementById("assessRisks").onclick = assessRisks;
  document.getElementById("exportDataToSheet").onclick = exportDataToSheet;
  document.getElementById("exportDataToJSON").onclick = exportDataToJSON;
  document.getElementById("exportDataWithDate").onclick = exportDataWithDate;
  document.getElementById("submitButton").onclick = handleSubmit;
  document.getElementById("aggregate").onclick = aggregate;
  document.getElementById("aggregateExpenses").onclick = aggregateExpenses;
  document.getElementById("getData").onclick = getData;



}

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function validate(type) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values"]);
      await context.sync();

      range.values.forEach((row, i) => {
        row.forEach((value, j) => {
          const isValid = checkIsValid(value, type);
          const message = isValid ? `Valid ${type}` : `Invalid ${type}`;
          const currentCell = range.getCell(i, j);
          const messageCell = currentCell.getOffsetRange(0, 1);
          messageCell.values = [[message]];
          currentCell.format.font.color = isValid ? "black" : "red";
          messageCell.format.font.color = isValid ? "black" : "red";
        });
      });

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function createDashboard() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.add("A1:D1", true);
      table.name = "PortfolioDashboard";
      table.getHeaderRowRange().values = [["Metric", "Value"]];
      table.rows.add(null, [
        ["Total Investments", "=SUM(Investments[Amount])"],
        ["Average ROI", "=AVERAGE(Investments[ROI])"],
        ["Total Valuation", "=SUM(Investments[Current Valuation])"]
      ]);
      table.columns.getItemAt(1).getRange().format.fill.color = "#FFFF00";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function trackKPIs() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.add("A1:E1", true);
      table.name = "KPITracker";
      table.getHeaderRowRange().values = [["Company", "Revenue", "CAC", "LTV", "Burn Rate"]];
      table.rows.add(null, [
        ["Company A", 500000, 5000, 15000, 20000],
        ["Company B", 1000000, 10000, 30000, 50000]
      ]);
      table.columns.getItemAt(1).getRange().format.fill.color = "#00FF00";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function assessRisks() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.add("A1:D1", true);
      table.name = "RiskAssessment";
      table.getHeaderRowRange().values = [["Investment", "Market Risk", "Execution Risk", "Financial Risk"]];
      table.rows.add(null, [
        ["Investment A", "High", "Medium", "Low"],
        ["Investment B", "Medium", "High", "High"]
      ]);
      table.columns.getItemAt(1).getRange().format.fill.color = "#FF0000";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function exportDataToSheet() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sourceSheet = workbook.worksheets.getActiveWorksheet();
      const newSheet = workbook.worksheets.add("Exported Data to Sheet");

      // Read data from the existing sheet
      const range = sourceSheet.getUsedRange();
      range.load(["values", "address"]);
      await context.sync();

      const data = range.values;

      // Add headers to the new sheet
      newSheet.getRange("A1:C1").values = [["Metric", "Value", "Notes"]];


      // Populate data
      newSheet.getRange(`A2:C${data.length + 1}`).values = data.map(row => [
        row[0], // Metric
        row[1], // Value
        "Notes here" // Notes (can be customized as needed)
      ]);


      // Format the header
      const header = newSheet.getRange("A1:C1");
      header.format.fill.color = "#4472C4";
      header.format.font.color = "white";
      header.format.font.bold = true;

      // Auto-fit columns for better readability
      newSheet.getRange("A:C").format.autofitColumns();

      await context.sync();
      console.log("Data exported successfully.");
    });
  } catch (error) {
    console.error(error);
  }
}

// export data to JSON
export async function exportDataToJSON() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sourceSheet = workbook.worksheets.getActiveWorksheet();

      // Read data from the existing sheet
      const range = sourceSheet.getUsedRange();
      range.load(["values", "text", "address"]);
      await context.sync();

      const data = range.values;
      const textData = range.text;

      // Create JSON data with type information
      const jsonData = data.map((row, rowIndex) => 
        row.map((cell, colIndex) => ({
          metric: cell, // Value
          type: typeof cell, // Type
          displayText: textData[rowIndex][colIndex] // Display text (for formatted values)
        }))
      );

      // Send JSON data to the API
      // const response = await fetch('API-ENDPOINT !!! TODO'), {
      //   method: 'POST',
      //   headers: {
      //     'Content-Type': 'application/json'
      //   },
      //   body: JSON.stringify(jsonData)
      // });

      // if (!response.ok) {
      //   throw new Error('Network response was not ok');
      // }

      // console.log('Data sent to the API successfully.');

      // Save JSON data to a file
      const jsonString = JSON.stringify(jsonData, null, 2);
      const blob = new Blob([jsonString], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'exportedData.json';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      // Optional: Add headers to a new sheet and populate data for user review
      const newSheet = workbook.worksheets.add("Exported Data");
      newSheet.getRange("A1:D1").values = [["Metric", "Type", "Display Text", "Notes"]];
      newSheet.getRange(`A2:D${data.length + 1}`).values = jsonData.map(row => 
        row.map(cell => [cell.metric, cell.type, cell.displayText, "Notes here"])
      );

      // Format the header
      const header = newSheet.getRange("A1:D1");
      header.format.fill.color = "#4472C4";
      header.format.font.color = "white";
      header.format.font.bold = true;

      // Auto-fit columns for better readability
      newSheet.getRange("A:D").format.autofitColumns();

      await context.sync();
      console.log("Data exported successfully.");
    });
  } catch (error) {
    console.error(error);
  }
}

export async function exportDataWithDate() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sourceSheet = workbook.worksheets.getActiveWorksheet();
      const logSheet = workbook.worksheets.getItem("ChangeLog");

      // Read data from the existing sheet
      const range = sourceSheet.getUsedRange();
      range.load(["values", "text", "address", "format/hidden"]);
      await context.sync();

      const data = range.values;
      const textData = range.text;
      const hiddenData = range.format.hidden;

      // Read data from the ChangeLog sheet
      const logRange = logSheet.getUsedRange();
      logRange.load("values");
      await context.sync();

      const logData = logRange.values;

      // Create a dictionary to store the last modified time for each cell
      const modifiedTimes = {};
      logData.slice(1).forEach(row => {
        const [timestamp, cellAddress] = row;
        modifiedTimes[cellAddress] = timestamp;
      });

      // Function to get the last modified time
      const getLastModifiedTime = (cellAddress) => {
        return modifiedTimes[cellAddress] || "Unknown";
      };

      // Create JSON data with type information, visibility status, and modification time
      const jsonData = data.map((row, rowIndex) =>
        row.map((cell, colIndex) => ({
          metric: cell, // Value
          type: typeof cell, // Type
          displayText: textData[rowIndex][colIndex], // Display text (for formatted values)
          hidden: hiddenData[rowIndex][colIndex], // Visibility status
          // timeModified: getLastModifiedTime(range.getCell(rowIndex, colIndex).address) // Last modification time
        }))
      );

      // Save JSON data to a file
      const jsonString = JSON.stringify(jsonData, null, 2);
      const blob = new Blob([jsonString], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'exportedData.json';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      // Optional: Add headers to a new sheet and populate data for user review
      const newSheet = workbook.worksheets.add("Exported Data with Date");
      // newSheet.getRange("A1:F1").values = [["Metric", "Type", "Display Text", "Hidden", "Time Modified", "Notes"]];
      newSheet.getRange("A1:E1").values = [["Metric", "Type", "Display Text", "Hidden", "Notes"]];

      newSheet.getRange(`A2:E${data.length + 1}`).values = jsonData.flatMap(row =>
        // row.map(cell => [cell.metric, cell.type, cell.displayText, cell.hidden, cell.timeModified, "Notes here"])
        row.map(cell => [cell.metric, cell.type, cell.displayText, cell.hidden, "Notes here"])

      );

      // Format the header
      const header = newSheet.getRange("A1:E1");
      header.format.fill.color = "#4472C4";
      header.format.font.color = "white";
      header.format.font.bold = true;

      // Auto-fit columns for better readability
      newSheet.getRange("A:E").format.autofitColumns();

      await context.sync();
      console.log("Data exported successfully.");
    });
  } catch (error) {
    console.error(error);
  }
}


// src/excelFunctions.js
export async function getData() {
  try {
    // Sample URL for the API endpoint (replace with your actual endpoint)
    // const url = 'https://your-api-endpoint.com/get-due-diligence-questions';

    // Fetch the data from the API
    // const response = await fetch(url);
    // if (!response.ok) {
    //   throw new Error('Network response was not ok');
    // }
    // const data = await response.json();

    // Dummy data for testing
    const data = [
      {
        question: "What is the company's revenue?",
        answer: "1 million USD",
        type: "Financial",
        lastReviewed: "2024-01-01",
        previousQuarterResponse: "900k USD"
      },
      {
        question: "Who are the key stakeholders?",
        answer: "John Doe, Jane Smith",
        type: "Operational",
        lastReviewed: "2024-01-02",
        previousQuarterResponse: "John Doe, Jane Smith"
      },
      {
        question: "What is the market risk level?",
        answer: "Medium",
        type: "Risk Assessment",
        lastReviewed: "2024-01-03",
        previousQuarterResponse: "High"
      }
    ];

    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.add("Due Diligence");

      // Define headers including "Previous Quarter's Response"
      const headers = [["Question", "Answer", "Type", "Last Reviewed", "Previous Quarter's Response"]];
      const dummyData = data.map(item => [
        item.question,
        item.answer || "",
        item.type,
        item.lastReviewed || "",
        item.previousQuarterResponse || ""
      ]);

      // Combine headers and data
      const combinedData = headers.concat(dummyData);

      // Write data to the sheet
      const range = sheet.getRangeByIndexes(0, 0, combinedData.length, headers[0].length);
      range.values = combinedData;

      // Format the header
      const headerRange = sheet.getRange("A1:E1");
      headerRange.format.fill.color = "#4472C4";
      headerRange.format.font.color = "white";
      headerRange.format.font.bold = true;

      // Auto-fit columns for better readability
      sheet.getUsedRange().format.autofitColumns();

      await context.sync();
      console.log("Data imported successfully.");
    });
  } catch (error) {
    console.error(error);
  }
}



function convertSerialNumberToDate(serialNumber) {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // Excel's epoch date
  const date = new Date(excelEpoch.getTime() + (serialNumber * 86400000)); // Convert to JavaScript date
  return date.toISOString().split('T')[0]; // Format as yyyy-mm-dd
}

// export async function handleSubmit() {
//   try {
//     await Excel.run(async (context) => {
//       const workbook = context.workbook;

//       // Check if "Form Data" sheet exists, create it if it doesn't
//       let sheet;
//       try {
//         sheet = workbook.worksheets.getItem("Form Data");
//         await context.sync(); // Ensure the sheet is loaded
//       } catch (error) {
//         sheet = workbook.worksheets.add("Form Data");
//         await context.sync(); // Ensure the new sheet is created
//       }

//       // Get form values
//       const metric = document.getElementById("metric").value;
//       const value = parseFloat(document.getElementById("value").value);
//       const notes = document.getElementById("notes").value;
//       const isApproved = document.getElementById("isApproved").checked;

//       // Find the next empty row in the "Form Data" sheet
//       let range;
//       try {
//         range = sheet.getUsedRange();
//         range.load("rowCount");
//         await context.sync();
//       } catch (error) {
//         // If there is no used range, start at row 1
//         range = { rowCount: 0 };
//       }

//       const nextRow = range.rowCount + 1;
//       const rowAddress = `A${nextRow}`;

//       // Insert form values into the sheet
//       sheet.getRange(`A${nextRow}:D${nextRow}`).values = [[metric, value, notes, isApproved]];

//       // Format header if first row
//       if (nextRow === 2) {
//         sheet.getRange("A1:D1").values = [["Metric", "Value", "Notes", "Is Approved"]];
//         const header = sheet.getRange("A1:D1");
//         header.format.fill.color = "#4472C4";
//         header.format.font.color = "white";
//         header.format.font.bold = true;
//       }

//       await context.sync();
//       console.log("Data submitted successfully.");
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

export async function handleSubmit() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;

      // Check if "Form Data" sheet exists, create it if it doesn't
      let sheet;
      try {
        sheet = workbook.worksheets.getItem("Form Data");
        await context.sync(); // Ensure the sheet is loaded
      } catch (error) {
        sheet = workbook.worksheets.add("Form Data");
        await context.sync(); // Ensure the new sheet is created
      }

      // Get form values
      const metric = document.getElementById("metric").value;
      const value = parseFloat(document.getElementById("value").value);
      const notes = document.getElementById("notes").value;
      const isApproved = document.getElementById("isApproved").checked;

      // Find the next empty row in the "Form Data" sheet
      let range;
      try {
        range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();
      } catch (error) {
        // If there is no used range, start at row 1
        range = { rowCount: 0 };
      }

      const nextRow = range.rowCount + 1;
      const rowAddress = `A${nextRow}`;

      // Insert form values into the sheet
      sheet.getRange(`A${nextRow}:D${nextRow}`).values = [[metric, value, notes, isApproved]];

      // Format header if first row
      if (nextRow === 2) {
        sheet.getRange("A1:D1").values = [["Metric", "Value", "Notes", "Is Approved"]];
        const header = sheet.getRange("A1:D1");
        header.format.fill.color = "#4472C4";
        header.format.font.color = "white";
        header.format.font.bold = true;
      }

      // Shade every other row light blue
      const usedRange = sheet.getUsedRange();
      usedRange.load("rowCount");
      await context.sync();

      for (let i = 1; i < usedRange.rowCount; i++) {
        if (i % 2 === 1) { // Every other row (odd rows, considering header at row 1)
          const rowRange = sheet.getRange(`A${i + 1}:D${i + 1}`);
          rowRange.format.fill.color = "#ADD8E6"; // Light blue color
        }
      }

      await context.sync();
      console.log("Data submitted and rows shaded successfully.");
    });
  } catch (error) {
    console.error(error);
  }
}



export async function aggregate() {
  try {
    await Excel.run(async (context) => {
      console.log("Reached aggregate recurring revenue");

      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);


      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();
      // const range = sheet.getSelectedRange();
      range.load("values");
      await context.sync();

      const data = range.values;
      console.log(data);

      let totalRecurringRevenue = 0;
      let isValidData = true;

      // Parse the data to find and aggregate recurring revenue values
      data.forEach(row => {
        row.forEach(cell => {
          if (typeof cell === 'number') {
            totalRecurringRevenue += cell;
          } else {
            isValidData = false;
          }
        });
      });

      if (!isValidData) {
        console.error("The selected range contains non-numeric values. Please select a range with numeric values only.");
        alert("The selected range contains non-numeric values. Please select a range with numeric values only.");
        return;
      }

      console.log(`Total Recurring Revenue: ${totalRecurringRevenue}`);

      // Create a new sheet to display the result
      const newSheet = workbook.worksheets.add("Aggregated Recurring Revenue");
      newSheet.activate();

      // Set the result in the new sheet
      const resultRange = newSheet.getRange("A1:B1");
      resultRange.values = [["Metric", "Value"]];
      const valueRange = newSheet.getRange("A2:B2");
      valueRange.values = [["Total Recurring Revenue", totalRecurringRevenue]];

      // Format the header and the result
      resultRange.format.fill.color = "#4472C4";
      resultRange.format.font.color = "white";
      resultRange.format.font.bold = true;
      valueRange.format.fill.color = "#4472C4";
      valueRange.format.font.color = "white";
      valueRange.format.font.bold = true;

      await context.sync();
      console.log("Aggregated data displayed successfully in the new sheet.");
    });
  } catch (error) {
    console.error(error);
  }
}

export async function aggregateExpenses() {
  try {
    await Excel.run(async (context) => {
      console.log("Reached aggregate expenses");

      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowIndex", "columnIndex", "values"]);
      await context.sync();
      console.log(`The range address was ${range.address}.`);

      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();

      const data = range.values;
      console.log(data);

      let totalExpenses = 0;
      let isValidData = true;

      // Check if the cell includes the word "expense" and add up the values 4 columns away
      for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
        for (let colIndex = 0; colIndex < data[rowIndex].length; colIndex++) {
          const cell = data[rowIndex][colIndex];
          if (typeof cell === 'string' && cell.toLowerCase().includes("expense")) {
            const targetCell = sheet.getRangeByIndexes(range.rowIndex + rowIndex, range.columnIndex + colIndex + 4, 1, 1);
            targetCell.load("values");
            await context.sync();

            const targetValue = targetCell.values[0][0];
            if (typeof targetValue === 'number') {
              totalExpenses += targetValue;
            } else {
              isValidData = false;
            }
          }
        }
      }

      if (!isValidData) {
        console.error("The selected range contains non-numeric values in the target cells. Please ensure the target cells contain numeric values.");
        alert("The selected range contains non-numeric values in the target cells. Please ensure the target cells contain numeric values.");
        return;
      }

      console.log(`Total Expenses: ${totalExpenses}`);

      // Create a new sheet to display the result
      const newSheet = workbook.worksheets.add("Aggregated Expenses");
      newSheet.activate();

      // Set the result in the new sheet
      const resultRange = newSheet.getRange("A1:B1");
      resultRange.values = [["Metric", "Value"]];
      const valueRange = newSheet.getRange("A2:B2");
      valueRange.values = [["Total Expenses", totalExpenses]];

      // Format the header and the result
      resultRange.format.fill.color = "#4472C4";
      resultRange.format.font.color = "white";
      resultRange.format.font.bold = true;
      valueRange.format.fill.color = "#4472C4";
      valueRange.format.font.color = "white";
      valueRange.format.font.bold = true;

      await context.sync();
      console.log("Aggregated data displayed successfully in the new sheet.");
    });
  } catch (error) {
    console.error(error);
  }
}


// export async function aggregateExpenses() {
//   try {
//     await Excel.run(async (context) => {
//       console.log("Reached aggregate expenses");

//       const range = context.workbook.getSelectedRange();
//       range.load("address");
//       // range.format.fill.color = "yellow";
//       await context.sync();
//       console.log(`The range address was ${range.address}.`);

//       const workbook = context.workbook;
//       const sheet = workbook.worksheets.getActiveWorksheet();
//       range.load("values");
//       await context.sync();

//       const data = range.values;
//       console.log(data);

//       let totalRecurringRevenue = 0;
//       let isValidData = true;

//       // Parse the data to find and aggregate recurring revenue values
//       data.forEach(row => {
//         row.forEach(cell => {
//           if (typeof cell === 'string') {

//             if (cell.includes("expenses")) {
//               range.format.fill.color = "magenta";

//             }

//             // totalRecurringRevenue += cell;

//           } else {
//             isValidData = false;
//           }
//         });
//       });

//       if (!isValidData) {
//         console.error("The selected range contains non-numeric values. Please select a range with numeric values only.");
//         alert("The selected range contains non-numeric values. Please select a range with numeric values only.");
//         return;
//       }

//       console.log(`Total Recurring Revenue: ${totalRecurringRevenue}`);

//       // Create a new sheet to display the result
//       const newSheet = workbook.worksheets.add("Aggregated Recurring Revenue");
//       newSheet.activate();

//       // Set the result in the new sheet
//       const resultRange = newSheet.getRange("A1:B1");
//       resultRange.values = [["Metric", "Value"]];
//       const valueRange = newSheet.getRange("A2:B2");
//       valueRange.values = [["Total Recurring Revenue", totalRecurringRevenue]];

//       // Format the header and the result
//       resultRange.format.fill.color = "#4472C4";
//       resultRange.format.font.color = "white";
//       resultRange.format.font.bold = true;
//       valueRange.format.fill.color = "#4472C4";
//       valueRange.format.font.color = "white";
//       valueRange.format.font.bold = true;

//       await context.sync();
//       console.log("Aggregated data displayed successfully in the new sheet.");
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }


function checkIsValid(value, type) {
  switch (type) {
    case 'number':
      return typeof value === 'number' && !isNaN(value);
    case 'date':
      return value instanceof Date && !isNaN(value.getTime());
    default:
      return typeof value === type;
  }
}

function checkIsFilled(value) {
  return value !== null && value !== undefined && value !== '';
}
