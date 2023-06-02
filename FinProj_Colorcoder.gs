function formatCells() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (sheetName === "Output" || sheetName === "Directions") {
      return; // Skip the "Output" and "Direction" sheets
    }
    var range = sheet.getDataRange();
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();
    var formulas = range.getFormulas();
    var numberFormats = range.getNumberFormats(); // Get number formats
    var fontColors = range.getFontColors(); // Get font colors
    var fontFamilies = range.getFontFamilies(); // Get font families
    
    // Continue with the rest of the processing
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        // Skip if the cell is a date, time, or duration
        if (numberFormats[i][j].indexOf("d") !== -1 || numberFormats[i][j].indexOf("y") !== -1 ||
            numberFormats[i][j].indexOf("h") !== -1 || numberFormats[i][j].indexOf("m") !== -1 ||
            numberFormats[i][j].indexOf("s") !== -1) {
          continue;
        }

        if (values[i][j] !== "" && formulas[i][j] === "" && typeof values[i][j] === "number") {
          backgrounds[i][j] = "#d9ead3";
          // Set numeric values to no decimal places
          numberFormats[i][j] = "0";
        }
        
        if (typeof values[i][j] === "number" && (numberFormats[i][j].indexOf("%") !== -1)) {
          numberFormats[i][j] = "0.0%"; // sets percentage format to have 1 decimal point
        }
        
        if (typeof values[i][j] === "number" && (numberFormats[i][j].indexOf("¤") !== -1)) {
          numberFormats[i][j] = "#,##0.0"; // Attempt to set currency format to have 1 decimal point
        }

        // If a formula references another sheet, change the font color to blue
        if (formulas[i][j].indexOf('!') !== -1 && formulas[i][j].indexOf(sheetName) === -1) {
          fontColors[i][j] = "#5b0f00";
        } else {
          fontColors[i][j] = "#000000"; // Set the rest of the cells to black
        }

        // Set all cells' font to Calibri
        fontFamilies[i][j] = "Calibri";
      }
    }

    // Set the new backgrounds, number formats, font colors, and font families
    range.setBackgrounds(backgrounds);
    range.setNumberFormats(numberFormats);
    range.setFontColors(fontColors);
    range.setFontFamilies(fontFamilies);
  });
}
