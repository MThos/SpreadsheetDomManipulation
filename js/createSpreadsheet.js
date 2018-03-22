/*
 * @program:	    createSpreadsheet.js
 * @description:    creates the spreadsheet for index.html and manipulates it.
 * @author:         Mykel Agathos
 * @date:           Dec 14, 2017
 * @revision:	    v1.0
 *
 * Copyright © 2017
 */


// global variables
var changed = null; // check for changed cell for border edit
var cellToEdit = null; // cell currently being changed
var filterSet = null; // filter of allowed char values
var selectedRow = 1; // default to row 1
var selectedCol = 1; // default to col 1
var rows = 20; // default rows for table
var columns = 10; // default columns for table
var tblArray = []; // array to hold all cell data
var alphaNums = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789():="; // allowed chars
var isIE11 = ((window.navigator.userAgent).indexOf("Trident") !== -1);

/*
 * @function:       createSpreadsheet()
 * @description:    set the number of rows/columns and call buildTable
 * @param:          none
 * @returns:        none
 */
function createSpreadsheet() {
    // create 2d array to hold cell data
    for (var i = 0; i < rows; i++) {
        tblArray[i] = [];
        for (var j = 0; j < columns; j++) {
            tblArray[i][j] = "";
        }
    }
    // build table with above row/column values
    document.getElementById("spreadsheet").innerHTML = buildTable(rows, columns)
}

/*
 * @function:       buildTable()
 * @description:    build the spreadsheet table based on passed in row/columns.
 * @param:          rows, columns
 * @returns:        divHTML
 */
function buildTable(rows, columns) {
    var divHTML = "<table border='1' cellpadding='0' cellspacing='0' class='spreadsheetTable'>";

    divHTML += "<tr><th></th>";

    for (var j = 0; j < columns; j++)
        divHTML += "<th>" + String.fromCharCode(j + 65) + "</th>";

    divHTML += "</tr>";

    for (var i = 1; i <= rows; i++) {
        divHTML += "<tr>";
        divHTML += "<td id='" + i + "_0' class='baseColumn'>" + i + "</td>";

        for (j = 1; j <= columns; j++)
            divHTML += "<td id='" + i + "_" + j + "' class='alphaColumn' onclick='clickCell(this)'></td>";

        divHTML += "</tr>";
    }

    divHTML += "</table>";
    return divHTML;
}

/*
 * @function:       clickCell()
 * @description:    set row/column selected and change background color
 * @param:          ref
 * @returns:        none
 */
function clickCell(ref) {
    var rcArray = ref.id.split('_');
    selectedRow = rcArray[0];
    selectedCol = rcArray[1];

    ref.style.borderColor = "#fff";
    if (changed != null) {
        changed.style.borderColor = "";
    }
    changed = ref;
}

/*
 * @function:       saveData()
 * @description:    convert tblArray to json and save to local storage
 * @param:          none
 * @returns:        none
 */
function saveData() {
    var tblJson = JSON.stringify(tblArray);
    localStorage.setItem("data", tblJson);
}

/*
 * @function:       loadData()
 * @description:    clear table then convert json from local storage
 *                  to tblArray, loop through the array and insert cell by
 *                  cell into the spreadsheet
 * @param:          none
 * @returns:        none
 */
function loadData() {
    clearTable();
    // grab the data from local storage and put back in tblArray
    var retrievedData = localStorage.getItem("data");
    tblArray = [];
    tblArray = JSON.parse(retrievedData);
    // loop through the array retrieved from local storage
    for (var i = 0; i < rows; i++) {
        for (var j = 0; j < columns; j++) {
            var rowEdit = i+1; // offset row number
            var colEdit = j+1; // offset col header
            // push array value to row/col number of cell
            cellToEdit = document.getElementById(rowEdit + "_" + colEdit);
            cellToEdit.innerHTML = tblArray[i][j];
        }
    }

    recalculate(); // calculate any =SUM in the cells
}

/*
 * @function:       clearTable()
 * @description:    recreate empty table, empty text box and clear storage
 * @param:          none
 * @returns:        none
 */
function clearTable() {
    document.getElementById("spreadsheet").innerHTML = "";
    document.getElementById("formulaBox").value = "";
    //localStorage.clear();
    createSpreadsheet();
    selectFocus();
}

/*
 * @function:       selectFocus()
 * @description:    sets the focus on the formula box
 * @param:          none
 * @returns:        none
 */
function selectFocus() {
    document.getElementById("formulaBox").value = "";
    document.getElementById("formulaBox").focus();
    document.getElementById("formulaBox").select();
}

/*
 * @function:       filterText()
 * @description:    filter the text in formula box on 'enter' and place the
 *                  text in the cell that is currently selected
 * @param:          ref
 * @returns:        none
 */
function filterText(ref) {
    var formulaBoxVal = null;

    if (ref.id === "formulaBox") {
        filterSet = alphaNums; // set the filter from global var 'alphaNums'
    }

    // check if IE11 or Chrome/Edge/Safari and fill the selected cell
    if (isIE11) {
        if (window.event.keyCode === 13) {
            formulaBoxVal = document.getElementById("formulaBox").value;
            cellToEdit = document.getElementById(selectedRow + "_" + selectedCol);
            cellToEdit.innerHTML = formulaBoxVal;
            tblArray[selectedRow-1][selectedCol-1] = formulaBoxVal;
            calculateCell(selectedRow-1, selectedCol-1);
        }
        else if (!nCharOK(window.event.keyCode)) {
            // prevent non-approved chars
            window.event.preventDefault();
        }

    }
    else {
        if (window.event.keyCode === 13) {
            formulaBoxVal = document.getElementById("formulaBox").value;
            cellToEdit = document.getElementById(selectedRow + "_" + selectedCol);
            cellToEdit.innerHTML = formulaBoxVal;
            tblArray[selectedRow-1][selectedCol-1] = formulaBoxVal;
            calculateCell(selectedRow-1, selectedCol-1);
        }
        else if (!nCharOK(window.event.keyCode)) {
            // prevent non-approved chars
            window.event.returnValue = null;
        }
    }
}

/*
 * @function:       nCharOK()
 * @description:    check the character entered to see if it is allowed
 * @param:          char
 * @returns:        none
 */
function nCharOK(char) {
    var ch = (String.fromCharCode(char));
    ch = ch.toUpperCase();

    if (filterSet.indexOf(ch) !== -1) {
        return true;
    }
}

/*
 * @function:       getFormula()
 * @description:    check if entered value is a formula =SUM(Ay:Bx) and
 *                  returns an array with the formula contents
 * @param:          tblVal
 * @returns:        arr
 */
function getFormula(tblVal) {
    var pattern = /[:|\(|\)]/;
    var arr = tblVal.split(pattern);
    var sum = arr[0].toUpperCase();

    if (arr.length < 3) {
        return null;
    }
    else if (sum !== "=SUM") {
        return null;
    }
    else {
        return arr;
    }
}

/*
 * @function:       recalculate()
 * @description:    loops through the tblArray and calls calculateCell to
 *                  determine the =SUM value
 * @param:          none
 * @returns:        none
 */
function recalculate()
{
    for (var i = 0; i < rows; i++){
        for (var j = 0; j < columns; j++){
            // check to see if table element is a formula
            if (tblArray[i][j].indexOf("=SUM") !== -1) {
                // calculate the formula for cell
                calculateCell(i, j);
            }
        }
    }
}

/*
 * @function:       calculateCell()
 * @description:    if formula is found =SUM(Ay:Bx), parse through it to find
 *                  the from and to cells (row/col) and perform calculation
 *                  by getting the values from the tblArray and call getFormula
 *                  to make the final calculation
 * @param:          row, column
 * @returns:        none
 */
function calculateCell(row, column)
{
    // begin by getting the formula parts
    var tokenArray = getFormula(tblArray[row][column]);

    // tokenArray[1] and tokenArray[2] contain the from and to references

    if (tokenArray != null) {
        var fromColumn = tokenArray[1].substr(0, 1);
        var fromRow = tokenArray[1].substr(1, tokenArray[1].length - 1);

        var toColumn = tokenArray[2].substr(0, 1);
        var toRow = tokenArray[2].substr(1, tokenArray[2].length - 1);

        // assign the actual row/column index values for the tblArray
        var fromRowIndex = parseFloat(fromRow) - 1;
        var fromColIndex = fromColumn.charCodeAt(0) - 65;

        var toRowIndex = parseFloat(toRow) - 1;
        var toColIndex = toColumn.charCodeAt(0) - 65;

        var fromVal = null;
        var toVal = null;

        // check to make sure values in cell are allowed
        if (isFloat(tblArray[fromRowIndex][fromColIndex])) {
            fromVal = tblArray[fromRowIndex][fromColIndex];
        }
        if (isFloat(tblArray[toRowIndex][toColIndex])) {
            toVal = tblArray[toRowIndex][toColIndex];
        }

        // add from and to values to the sumTotal
        var sumTotal = parseFloat(fromVal) + parseFloat(toVal);

        // insert sumTotal into spreadsheet cell (offset row/col)
        var cell = (row + 1) + "_" + (column + 1);
        var ref = document.getElementById(cell);
        ref.innerHTML = sumTotal;
    }
}

/*
 * @function:       isFloat()
 * @description:    check if float values are accepted
 * @param:          val
 * @returns:        true/false
 */
function isFloat(val)
{
    var ch = "";
    var justFloat = "0123456789.";

    for (var i = 0; i < val.length; i++) {
        ch = val.substr(i, 1);

        if (justFloat.indexOf(ch) === -1) {
            return false;
        }
    }
    return true;
}

/* Copyright © 2017 */