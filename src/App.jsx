import React, { useState } from "react";
import "./App.css";
import GC from "@mescius/spread-sheets";
import { SpreadSheets, Worksheet } from "@mescius/spread-sheets-react";
import * as ExcelIO from "@mescius/spread-excelio";
import saveAs from "file-saver";
import "@mescius/spread-sheets-charts";

// Your SpreadJS License Key
var SpreadJSKey =
  "UPES,E838759459758123#B13lVNJFmQYxUdrJUVFZndpplbMxmbzIlbuRjZYV6Rsd6QQZGcVVETMh7VQJkdDBFRZZnWZ3GO6YTMqZ6LGlDOxMXa83id5lTSmhWQ4MnWvE4Rs5WRzdUaB54cTNFTLdleEFGTslXd4VXbuVDZhdXctRjSXxERk3GM6I7MYh4T9AVS4YGTHRFUBVUUxM6ahJ5KQJTdEhXRV5ke5MHS6pncLFjYRtGMFJTUrFXSxlHbhZjZGhEUrQ5VCN5YVx6UxUmUFlHRWJGNQJWRCx6LIJTMM3Cc486KkZ4Z984MRxGezlTNzUXbLZ5TjdXUzQzcOJiOiMlIsISM5gjQBljQiojIIJCL7YjN6AzM6gzM0IicfJye35XX3JSVBtUUiojIDJCLicTMuYHITpEIkFWZyB7UiojIOJyebpjIkJHUiwiI4QTM4YDMgETM9ADNyAjMiojI4J7QiwiIxEDMxQjMwIjI0ICc8VkIsIyUFBVViojIh94QiwSZ5JHd0ICb6VkIsIyMyEDO5cTO5QTO5cDOzgjI0ICZJJCL355W0IyZsZmIsU6csFmZ0IiczRmI1pjIs9WQisnOiQkIsISP3cGNZRlUvIlUa5WehdDRTZkerRXTH9UQUBXel5UVq34QER6Y8g4Y4cWbq3idCFDbUd7Z5o5ZUJVaDRmYFRmdUJVdVVWYnlEW6oGRnh5V7I7dGRlV8wkZERjRtRzKwIGZTZjRPRGMddTN"; // Your actual license key
GC.Spread.Sheets.LicenseKey = SpreadJSKey;

if (ExcelIO.setLicenseKey) {
  ExcelIO.setLicenseKey(SpreadJSKey);
}

const App = () => {
  const [spread, setSpread] = useState(null);
  const hostStyle = {
    width: "1100px",
    height: "800px",
  };

  // /////////////////////////
  const initializeCharts = (spread) => {
    let sheetCharts = spread.getSheet(0);

    // Set the sheet name
    sheetCharts.name("Charts");
    sheetCharts.suspendPaint();

    // // Prepare data for chart
    // sheetCharts.setValue(0, 1, "Q1");
    // sheetCharts.setValue(0, 2, "Q2");
    // sheetCharts.setValue(0, 3, "Q3");
    // sheetCharts.setValue(1, 0, "Mobile Phones");
    // sheetCharts.setValue(2, 0, "Laptops");
    // sheetCharts.setValue(3, 0, "Tablets");

    // for (let r = 1; r <= 3; r++) {
    //   for (let c = 1; c <= 3; c++) {
    //     sheetCharts.setValue(r, c, parseInt(Math.random() * 100));
    //   }
    // }

    // Add columnClustered chart
    // let chart_columnClustered = sheetCharts.charts.add(
    //   "chart_columnClustered",
    //   GC.Spread.Sheets.Charts.ChartType.columnClustered,
    //   5,
    //   150,
    //   300,
    //   300,
    //   "A1:D4"
    // );
    // chart_columnClustered.title({ text: "Annual Sales" });

    // Add columnStacked chart
    // let chart_columnStacked = sheetCharts.charts.add(
    //   "chart_columnStacked",
    //   GC.Spread.Sheets.Charts.ChartType.columnStacked,
    //   320,
    //   150,
    //   300,
    //   300,
    //   "A1:D4"
    // );
    // chart_columnStacked.title({ text: "Annual Sales" });

    // // Add columnStacked100 chart
    // let chart_columnStacked100 = sheetCharts.charts.add(
    //   "chart_columnStacked100",
    //   GC.Spread.Sheets.Charts.ChartType.columnStacked100,
    //   640,
    //   150,
    //   300,
    //   300,
    //   "A1:D4"
    // );
    // chart_columnStacked100.title({ text: "Annual Sales" });

    // Resume paint to render everything
    sheetCharts.resumePaint();
  };

  // Handling workbook initialization event
  const workbookInit = (spreadInstance) => {
    setSpread(spreadInstance);
    initializeCharts(spreadInstance); // Initialize charts once the workbook is loaded
  };

  // Handling workbook initialized event
  // const workbookInit = (spread) => {
  //   setSpread(spread);
  // };

  // /////////////////

  // Import Excel with Charts
  const importFile = () => {
    const excelFile = document.getElementById("fileDemo").files[0];
    const excelIO = new ExcelIO.IO();

    // Open the file and handle the response with JSON
    excelIO.open(
      excelFile,
      (json) => {
        spread.fromJSON(json);

        // Force SpreadJS to re-render and ensure charts are loaded
        spread.refresh();
      },
      (e) => {
        console.error(e);
      }
    );
  };

  // Export Excel
  const exportFile = () => {
    const excelIO = new ExcelIO.IO();
    let fileName = document.getElementById("exportFileName").value;
    if (fileName.substr(-5, 5) !== ".xlsx") {
      fileName += ".xlsx";
    }
    const json = JSON.stringify(spread.toJSON());
    excelIO.save(
      json,
      (blob) => {
        saveAs(blob, fileName);
      },
      (e) => {
        console.error(e);
      }
    );
  };

  return (
    <div>
      <input type="file" name="files[]" id="fileDemo" accept=".xlsx" />
      <input type="button" id="loadExcel" value="Import" onClick={importFile} />
      <input type="button" id="saveExcel" value="Export" onClick={exportFile} />
      <input
        type="text"
        id="exportFileName"
        placeholder="Export file name"
        defaultValue="export.xlsx"
      />
      <SpreadSheets hostStyle={hostStyle} workbookInitialized={workbookInit}>
        <Worksheet />
      </SpreadSheets>
    </div>
  );
};

export default App;
