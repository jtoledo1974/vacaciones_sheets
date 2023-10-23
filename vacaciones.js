function testOnEdit() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var range = sheet.getRange("F8");

    // Create mock event object
    var e = {
        range: range,
        // Add more event object attributes as needed
    };

    onEdit(e);

    // Check checkNameFormat
    var name = "Juan (1)";
    var result = checkNameFormat(name);
    Logger.log(result);
    Logger.log(result.cell);
    Logger.log(result.order);
}

function editA1() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(33, 4).setValue(Math.random());
}


function onEdit(e) {
    var range = e.range;
    var sheet = SpreadsheetApp.getActiveSpreadsheet()

    // Return if multiple cells are selected.

    if (!isSingleCell(range)) {
        sheet.toast("Multiple cells selected. Please select only one cell.")
        return;
    }

    // Return if we are not filling up a cell 
    var date = getCycleFromRange(range)
    if (!date) {
        return;
    }

    // Return if the edit value does not pass checkNameFormat
    var name_order = checkNameFormat(range.getValue());
    if (!name_order) {
        return;
    }

    // Record the date in the appropriate cell
    recordDate(name_order.cell, name_order.order, date);

    // Record the points
    recordPoints(name_order.cell, name_order.order, range.getValue());

    // Check if the input value matches a name from the named range NombresPeticiones.
    // If it does, set the background color of the cell to green.
    // If it doesn't, set the background color of the cell to red.

    var name = range.getValue();
    var cell = getCellFromName(name);



    var row = range.getRow();
    var col = range.getColumn();
    var message = "Cell edited: Row " + row + ", Column " + col;
    sheet.toast(message);
}

function isSingleCell(range) {
    return range.getNumRows() === 1 && range.getNumColumns() === 1;
}


function recordPoints(nameCell, round, points) {
/**
 * Adds the points to the points area.
 * 
 * @param {*} nameCell - The first cell in the names range for the points are
 * @param {*} round - Ronda de petición
 * @param {*} date - Ciclo de vacaciones
 */
    Logger.log("recordPoints(" + nameCell + ", " + round + ", " + points + ")");
    // Make sure nameRow and nameCol are numbers
    var nameRow = parseInt(nameCell.getRow());
    var nameCol = parseInt(nameCell.getColumn());
    Logger.log("nameCell (row, col): " + nameCell.getValues() + " (" + nameRow + ", " + nameCol + ")");

    var roundCol = parseInt(nameCol + 2 * round);
    Logger.log("roundCol: " + roundCol)

    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet.toast("Setting cell (" + nameRow + ", " + roundCol + ") to " + points);

    sheet.getActiveSheet().getRange(nameRow, roundCol).setValue(points);
}


// Given a name cell, a round number, and a date, fill the appropriate cell with the date
function recordDate(nameCell, round, date) {
    Logger.log("recordDate(" + nameCell + ", " + round + ", " + date + ")");
    // Make sure nameRow and nameCol are numbers
    var nameRow = parseInt(nameCell.getRow());
    var nameCol = parseInt(nameCell.getColumn());
    Logger.log("nameCell (row, col): " + nameCell.getValues() + " (" + nameRow + ", " + nameCol + ")");

    var roundCol = parseInt(nameCol + round);
    Logger.log(roundCol)

    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet.toast("Setting cell (" + nameRow + ", " + roundCol + ") to " + date);

    sheet.getActiveSheet().getRange(nameRow, roundCol).setValue(date);
}


function checkNameFormat(value) {
/**
 * Check name format. Checks that the given value is in the format Name (n), 
 * where Name is a valid name as getCellFromName would return, and n is a number
 * between 1 and 6.
 * If everything checks, we return the cell from the NombresPeticiones range and the number
 * Otherwise, we return null
 */    
    Logger.log("Checking name format for " + value);
    var name = value.substring(0, value.indexOf("(") - 1);
    var order = value.substring(value.indexOf("(") + 1, value.indexOf(")"));
    if (order < 1 || order > 6) {
        return null;
    }
    var cell = getCellFromName(name);
    Logger.log("Cell coords: " + cell.getRow() + ", " + cell.getColumn() + " (" + cell.getValues() + ")");
    Logger.log("Number: " + order)
    if (cell) {
        return {
            cell: cell,
            order: parseInt(order)
        };
    }
    return null;
}

// Checks whether a string matches any of the values of the cells in the named range NombresPeticiones.
// We use getRangeByName.  If it is, return the cell that matches
function getCellFromName(name) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = sheet.getRangeByName("NombresPeticiones");
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
        if (values[i][0] === name) {
            return range.getCell(i + 1, 1);
        }
    }
    return null;
}



//Checks weather the cell either two, three or four cells above is contained in one of 
// the named ranges PreVerano, Verano or PosVerano
// If it is, return the value for that cell, otherwise null
function getCycleFromRange(range) {
    namedRanges = ["PreVerano", "Verano", "PosVerano"];
    for (var i = 0; i < namedRanges.length; i++) {
        if (cellInNamedRange(range.offset(-2, 0), namedRanges[i])) {
            // Return the value of the cell two cells above
            return range.offset(-2, 0).getValue();
        }
        if (cellInNamedRange(range.offset(-3, 0), namedRanges[i])) {
            // Return the value of the cell three cells above
            return range.offset(-3, 0).getValue();
        }
        if (cellInNamedRange(range.offset(-4, 0), namedRanges[i])) {
            // Return the value of the cell four cells above
            return range.offset(-4, 0).getValue();
        }
        return null;
    }

}

// Checks whether a given range is contained in a given named range
function cellInNamedRange(cell, rangeName) {
    var namedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
    return isCellInRange(cell, namedRange);
}

function isCellInRange(cell, range) {
    // Get the row and column indices for the top-left corner of the range
    var rangeRowStart = range.getRow();
    var rangeColStart = range.getColumn();

    // Get the row and column indices for the bottom-right corner of the range
    var rangeRowEnd = rangeRowStart + range.getNumRows() - 1;
    var rangeColEnd = rangeColStart + range.getNumColumns() - 1;

    // Get the row and column indices for the cell
    var cellRow = cell.getRow();
    var cellCol = cell.getColumn();

    // Check if the cell is within the range
    return (cellRow >= rangeRowStart && cellRow <= rangeRowEnd) &&
        (cellCol >= rangeColStart && cellCol <= rangeColEnd);
}