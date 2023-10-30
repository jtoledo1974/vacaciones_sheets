var spreadsheet
var sheet

function testOnEdit() {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet = spreadsheet.getActiveSheet();
    var range = spreadsheet.getRange("E16");

    // Create mock event object
    var e = {
        range: range,
        // Add more event object attributes as needed
    };

    onEdit(e);

}

function editA1() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(33, 4).setValue(Math.random());
}


function onEdit(e) {
    var range = e.range;
    if (!spreadsheet) {
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        sheet = spreadsheet.getActiveSheet();
    }

    // Return if multiple cells are selected.

    if (!isSingleCell(range)) {
        spreadsheet.toast("Multiple cells selected. Please select only one cell.")
        return;
    }

    // Return if we are not filling up a cell 
    var date = getCycleFromRange(range)
    if (!date) {
        return;
    }
    var points = getPointsFromRange(range);
    if (!points) {
        return;
    }

    if (range.getValue() == "") {
        // Clearing a cell
        clearRequest(e.oldValue);
        return;
    }

    // Adding a new petition
    recordRequest(range.getValue(), date, points);
}

function recordRequest(name_order_string, date, points) {
    /**
     * Records a new petition in the spreadsheet.
     * 
     * @param {*} name_order_string - String in the format Name (n), where Name is a valid name as getCellFromName would return, and n is a number between 1 and 6.
     * @param {*} date - Ciclo de vacaciones
     * @param {*} points - Puntos de la petición
     * @returns null if the name_order_string is not in the correct format, the name and the order otherwise
     */

    // Return if the edit value does not pass checkNameFormat
    var name_order = checkNameFormat(name_order_string);
    if (!name_order) {
        return null;
    }


    // Record the date in the appropriate cell
    var cell = getCellFromNameAndRange(name_order.name, "NombresPeticiones");
    recordDate(cell, name_order.order, date);

    // Record the points
    cell = getCellFromNameAndRange(name_order.name, "NombresPuntos");
    recordPoints(cell, name_order.order, points);

    return name_order;

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

    spreadsheet.toast("Setting cell (" + nameRow + ", " + roundCol + ") to " + points);
    sheet.getRange(nameRow, roundCol).setValue(points);
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

    spreadsheet.toast("Setting cell (" + nameRow + ", " + roundCol + ") to " + date);
    sheet.getRange(nameRow, roundCol).setValue(date);
}


function checkNameFormat(value) {
    /**
     * Check name format. Checks that the given value is in the format Name (n), 
     * where Name is a valid name as getCellFromName would return, and n is a number
     * between 1 and 6.
     * 
     * @param {*} value - String to check for format Name (n)
     * @returns - The name and the order if the format is correct, null otherwise
     */
    Logger.log("checkNameFormat(" + value + ")");
    var name = value.substring(0, value.indexOf("(") - 1);
    var order = value.substring(value.indexOf("(") + 1, value.indexOf(")"));
    if (order < 1 || order > 6) {
        return null;
    }
    var cell = getCellFromNameAndRange(name, "NombresPeticiones");
    if (!cell) {
        Logger.log("checkNameFormat(" + value + ") -> null")
        return null;
    }
    Logger.log("checkNameFormat(" + value + ") -> {name: " + name + ", order: " + order + "}")
    return {
        name: name,
        order: parseInt(order)
    };

}

function getCellFromNameAndRange(name, namedRange) {
    /**
     * Checks whether a string matches any of the values of the cells in the given named range
     * 
     * @param {*} name - String to check for
     * @param {*} namedRange - Named range to check for the string
     * @returns - The cell that contains the string, null otherwise
     */
    Logger.log("getCellFromNameAndRange(" + name + ", " + namedRange + ")");
    var range = spreadsheet.getRangeByName(namedRange);
    var values = range.getValues();

    for (var i = 0; i < values.length; i++) {
        if (values[i][0] === name) {
            Logger.log("getCellFromNameAndRange(" + name + ", " + namedRange + ") -> range.getCell(" + (i + 1) + ", 1)");
            return range.getCell(i + 1, 1);
        }
    }
    Logger.log("getCellFromNameAndRange(" + name + ", " + namedRange + ") -> null");
    return null;
}



//Checks weather the cell either two, three or four cells above is contained in one of 
// the named ranges PreVerano, Verano or PosVerano
// If it is, return the value for that cell, otherwise null
function getCycleFromRange(range) {
    Logger.log("getCycleFromRange(" + range + ")");
    namedRanges = ["PreVerano", "Verano", "PosVerano"];
    var res = -1;
    for (var i = 0; i < namedRanges.length; i++) {
        if (cellInNamedRange(range.offset(-2, 0), namedRanges[i])) {
            // Return the value of the cell two cells above
            res = range.offset(-2, 0).getValue();
        }
        else if (cellInNamedRange(range.offset(-3, 0), namedRanges[i])) {
            // Return the value of the cell three cells above
            res = range.offset(-3, 0).getValue();
        }
        else if (cellInNamedRange(range.offset(-4, 0), namedRanges[i])) {
            // Return the value of the cell four cells above
            res = range.offset(-4, 0).getValue();
        }
    }
    if (res == -1) {
        Logger.log("getCycleFromRange(" + range + ") -> null");
        return null;
    }

    Logger.log("getCycleFromRange(" + range + ") -> " + res);
    return res;

}

function getPointsFromRange(range) {
    /**
     * Checks whether the cell either two, three or four cells above is contained in one of
     * the named ranges PreVerano, Verano or PosVerano
     * If it is, return the date from the points from the cell above the named range,
     * which contains the points as int.
     * 
     * @param {*} range - Range to check for points
     * @returns - The points as int, null otherwise
     */
    Logger.log("getPointsFromRange(" + range + ")");
    namedRanges = ["PreVerano", "Verano", "PosVerano"];
    var res = -1;
    for (var i = 0; i < namedRanges.length; i++) {
        if (cellInNamedRange(range.offset(-2, 0), namedRanges[i])) {
            // Return the value of the cell two cells above
            res = range.offset(-3, 0).getValue();
        }
        else if (cellInNamedRange(range.offset(-3, 0), namedRanges[i])) {
            // Return the value of the cell three cells above
            res = range.offset(-4, 0).getValue();
        }
        else if (cellInNamedRange(range.offset(-4, 0), namedRanges[i])) {
            // Return the value of the cell four cells above
            res = range.offset(-5, 0).getValue();
        }
    }
    if (res == -1) {
        Logger.log("getPointsFromRange(" + range + ") -> null");
        return null;
    }

    Logger.log("getPointsFromRange(" + range + ") -> " + res);
    return parseInt(res);

}

// Checks whether a given range is contained in a given named range
function cellInNamedRange(cell, rangeName) {
    var namedRange = spreadsheet.getRangeByName(rangeName);
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


function fillEnairePorFechas() {
    /**
     * Fills the EnairePorFechas table with the format required by Enaire.
     * 
     * @returns - null
     */

    // We go through each of the dates for pre verano, verano and pos verano
    // We check to see which people have signed up for that date
    // The we get the official name from the NombresCompletos named range
    // And fill a table with 4 columns: date, first requester, second requester, third requester

    // We get the spreadsheet
    if (!spreadsheet) {
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
        sheet = spreadsheet.getActiveSheet();
    }

    // For each of the three named ranges
    var namedRanges = ["PreVerano", "Verano", "PosVerano"];

    // get the EnairePorFechas named range
    var enairePorFechas = spreadsheet.getRangeByName("EnairePorFechas");

    for (var i = 0; i < namedRanges.length; i++) {
        var season_dates_and_names = getSeasonDatesAndNames(namedRanges[i]);
        for (var j = 0; j < season_dates_and_names.length; j++) {
            var date = season_dates_and_names[j].date;
            var names = season_dates_and_names[j].names;
            // appendRowEnairePorFechas(date, names);
        }

    }
    return null;
}


function getSeasonDatesAndNames(namedRange) {
    /**
     * Gets the dates and names for a given named range.
     * 
     * @param {*} namedRange - Named range to get the dates and names from
     * @returns - An array of objects with the date and the names for that date
     */
    Logger.log("getSeasonDatesAndNames(" + namedRange + ")");
    var range = spreadsheet.getRangeByName(namedRange);
    var res = [];

    // The season named range is a row with dates in each column
    // The names are filled in the cells below the dates
    // We go through each column, and for each column we get the date
    // and the names below it

    // Get the row and column indices for the top-left corner of the range
    var rangeRowStart = range.getRow();
    var rangeColStart = range.getColumn();

    // Get the row and column indices for the bottom-right corner of the range
    var rangeRowEnd = rangeRowStart + range.getNumRows() - 1;
    var rangeColEnd = rangeColStart + range.getNumColumns() - 1;

    // Get the row and column indices for the cell
    var cellRow = rangeRowStart;
    var cellCol = rangeColStart;

    // Get the values for the range
    var values = range.getValues();

    var official_names_values = spreadsheet.getRangeByName("NombresCompletos").getValues();

    // For each column
    for (var i = 0; i < values[0].length; i++) {
        // Get the date
        var date = values[0][i];
        // Get the names
        var names = [];
        for (var j = 2; j < 5; j++) {
            // The names are actually outside the range, so we need to get the value from the cells
            var name_order = sheet.getRange(cellRow + j, cellCol + i).getValue();
            var name = getOfficialName(name_order);
            names.push(name);
            Logger.log("(" + (cellRow + j) + ", " + (cellCol + i) + "): " + name);
        }
        // Add the date and the names to the result
        Logger.log("getSeasonDatesAndNames(" + namedRange + ") -> {date: " + date + ", names: " + names + "}");
        res.push({
            date: date,
            names: names
        });
    }

    Logger.log("getSeasonDatesAndNames(" + namedRange + ") -> " + res);
    return res;
}

function getOfficialName(name_order_string) {
    /**
     * Gets the official name from a name_order_string.
     * 
     * @param {*} name_order_string - String in the format Name (n), where Name is a valid name as getCellFromName would return, and n is a number between 1 and 6.
     * @returns - The official name, null if the name_order_string is not in the correct format
     */
    Logger.log("getOfficialName(" + name_order_string + ")");
    var name_order = checkNameFormat(name_order_string);
    if (!name_order) {
        Logger.log("getOfficialName(" + name_order_string + ") -> null");
        return null;
    }
    var name = name_order.name;

}