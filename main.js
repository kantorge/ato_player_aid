// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

/*
 * Main functions, which are called by the Shopping sheet buttons
 */

// Clear all trading settings
function clearTradeSettings() {
  clearNamedRange('TradeSettings');
}

// Clear selected items for purchase
function clearCartContent() {
  clearNamedRange('CartContent');
}

// Commit trade settings
function recordTradeSettings() {
  // Read trade settings
  var tradeSettings = readNamedRangeValues('TradeSettings');

  // Read source values
  var resourceInventory = readNamedRangeValues('ResourceInventory');

  // Loop rows and adjust source. We assume, that the main inventory settings are in the same order as the trade settings
  for (var i=0; i<resourceInventory.length; i++) {
    resourceInventory[i][0] = resourceInventory[i][0] - tradeSettings[i][0] + tradeSettings[i][1];

    // Check if changes would result in invalid value
    if (resourceInventory[i][0] < 0) {
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        "Error",
        "Trading would result in negative quantity for one or more resources. Please review your settings. No changes were made.",
        ui.ButtonSet.OK
      );

      // Abort recording of trade settings
      return;
    }
  }

  // Apply calculated values to main inventory
  writeToNamedRange('ResourceInventory', resourceInventory);

  // Finally, remove trade settings
  clearTradeSettings();
}

// Commit purchase based on selected items
function commitPurchase() {
  // References to columns in the gear list
  const gearOwnershipColumnIndex = 6;
  const gearNameColumnIndex = 0;

  // Read cost
  var cost = readNamedRangeValues('PurchaseCost');

  // Apply cost and validate if resources are available (not including trade)
  var resourceInventory = readNamedRangeValues('ResourceInventory');
  for (var i=0; i<resourceInventory.length; i++) {
    resourceInventory[i][0] = resourceInventory[i][0] - cost[i][0];

    // Check if changes would result in invalid value
    if (resourceInventory[i][0] < 0) {
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        "Error",
        "Purchase would result in negative quantity for one or more resources. Please review your settings or record your planned trading. No changes were made.",
        ui.ButtonSet.OK);

      // Abort recording of purchase settings
      return;
    }
  }

  // Read items to be purchased
  var items = readNamedRangeValues('CartContent');

  // Record purchased items
  var gears = readNamedRangeValues('FullGearList');
  items[0].forEach(function(item) {
    // Skip empty items
    if (!item) {
      return;
    }

    // Update one item of type from not owned to owned
    var index = gears.findIndex(gear => gear[gearNameColumnIndex] === item && gear[gearOwnershipColumnIndex] !== "Y");
    if (index < 0) {
      // This is an error, actually
      return;
    }
    gears[index][gearOwnershipColumnIndex] = "Y";
  })

  // Write new resources back
  writeToNamedRange('ResourceInventory', resourceInventory);

  // Save purchased items array
  writeToNamedRange('FullGearList', gears);

  // Clear cart content
  clearCartContent();
}

/*
 * Helper functions
 */

// Get reference to sheet by its ID
function findSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().find(s => s.getSheetId() === id);
}

// Get reference to sheet by its name
function findSheetByName(name) {
  return SpreadsheetApp.getActive().getSheets().find(s => s.getSheetName() === name);
}

// Read the contents of a named range
function readNamedRangeValues(rangeName) {
  var range = SpreadsheetApp.getActive().getRangeByName(rangeName);
  return range.getValues();
}

// Dump array into a named range
function writeToNamedRange(rangeName, array) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(rangeName);
  range.setValues(array);
}

// Clear the contents of a named range. Optionally, ask for confirmation.
function clearNamedRange(rangeName, getConfirmation) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(rangeName);

  if (getConfirmation) {
    var ui = SpreadsheetApp.getUi();

    var response = ui.alert("Delete Confirmation", "Are you sure you want to clear the values?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.NO) {
      return;
    }
  }

  range.clearContent();
}
