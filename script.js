// Initialise constants
const spreadsheetContainer = document.querySelector("#spreadsheet");
const ROWS = 101;
const COLS = 101;
const spreadsheet = [];

var selectedCell;

// Cell class
class Cell {
  constructor(
    col,
    row,
    cellData,
    colLet,
    rowNum,
    cellID,
    isHeader,
    displayValue,
    isBold,
    isItalics,
    isUnderlined
  ) {
    // Basic information for cells
    this.col = col;
    this.row = row;
    this.cellData = cellData;
    this.colLet = colLet;
    this.rowNum = rowNum;
    this.cellID = cellID;
    this.isHeader = isHeader;
    this.displayValue = displayValue;
    this.isBold = isBold;
    this.isItalics = isItalics;
    this.isUnderlined = isUnderlined;
  }
}

initMatrix();

// Initialise matrix for spreadsheet
function initMatrix() {
  for (let i = 0; i < COLS; i++) {
    let row = [];

    for (let j = 0; j < ROWS; j++) {
      // Sets default values for each cell object
      let cellData = "0";
      let displayValue = "";
      let isHeader = false;

      // If statements to check if cell should be editable
      if (j == 0 && i != 0) {
        displayValue = i;
      }
      if (i == 0 && j != 0) {
        displayValue = numToLetter(j);
      }
      if (j == 0 || i == 0) {
        isHeader = true;
      }
      const cell = new Cell(
        j,
        i,
        cellData,
        numToLetter(j),
        i,
        numToLetter(j) + i,
        isHeader,
        displayValue,
        false,
        false,
        false
      );
      row.push(cell);
    }
    spreadsheet.push(row);
  }
  render();
  console.log("Spreadsheet", spreadsheet); // For easier testing with the DOM
}

// Initialise` cells
function initCell(cell) {
  const cellElem = document.createElement("input");
  cellElem.className = "cell";
  cellElem.id = cell.colLet + cell.rowNum;
  cellElem.value = cell.displayValue;
  if (cell.isHeader) {
    cellElem.classList.add("header");
  }
  cellElem.disabled = cell.isHeader;
  cellElem.onchange = (e) => handleOnChange(e.target.value, cell);
  cellElem.onclick = (e) => handleOnClick(cell);
  return cellElem;
}

// Convert col numbers to corresponding col letters
function numToLetter(num) {
  var result;
  // First 26 columns
  if (num <= 26) {
    result = getLetter(num);
  } else {
    // Remaining columns up to 2 letters e.g. AA
    if (num % 26 == 0) {
      result = getLetter(Math.floor(num / 26) - 1) + getLetter(26);
    } else {
      result = getLetter(Math.floor(num / 26)) + getLetter(num % 26);
    }
  }
  return result;
}

// Returns letter value for col title
function getLetter(num) {
  return String.fromCharCode("A".charCodeAt(0) + num - 1);
}

// Returns cell based on ID
function getCell(ID) {
  let result;
  // Finds matching cellID within matrix
  spreadsheet.forEach((row) => {
    row.forEach((cell) => {
      if (cell.cellID == ID) {
        result = cell;
      }
    });
  });
  return result;
}

// Render spreadsheet
function render() {
  console.log("Rendering Spreadsheet..."); // Added this to show that it is redrawing the grid

  // Clears div every time render() is called
  spreadsheetContainer.innerHTML = "";

  for (let i = 0; i < spreadsheet.length; i++) {
    const rowContainer = document.createElement("div");
    rowContainer.className = "row";

    for (let j = 0; j < spreadsheet[i].length; j++) {
      const cell = spreadsheet[i][j];
      rowContainer.append(initCell(cell));
    }
    spreadsheetContainer.append(rowContainer);
  }
}

// Changes cell data
function handleOnChange(cellData, cell) {
  // toUpperCase() for input consistency
  cellData.toUpperCase();
  // Check if cell contains function, then equation, otherwise just update display value
  if (cellData.startsWith("=SUM") || cellData.startsWith("=AVERAGE")) {
    handleFunction(cellData.slice(1), cell);
  } else if (cellData.startsWith("=")) {
    parseEquation(cellData.slice(1), cell);
  } else {
    cell.cellData = cellData;
    updateDisplayValue(cellData, cell);
    updateReferences(cellData, cell);
  }
}

// Handles onclick event for cells
function handleOnClick(cell) {
  selectedCell = cell;
}

// Parse through the entire string
function parseEquation(cellData, cell) {
  let cells = [];
  let operators = [];
  let c = "";
  cell.cellData = cellData;

  for (let i = 0; i < cellData.length; i++) {
    // Checks if current char is operator, then if parser has reached end of line, then alphanumeric
    if (!cellData[i].match(/^[a-zA-Z0-9]+$/)) {
      // Pushes cell 'c' to the array of cells, as well as current operator
      cells.push(c);
      operators.push(cellData[i]);
      c = "";
    } else if (i == cellData.length - 1) {
      // Sets and pushes cell 'c' to the array of cells
      c += cellData[i];
      cells.push(c);
      c = "";
    } else {
      // Sets cell 'c'
      c += cellData[i];
    }
  }

  let result = getCell(cells[0]).cellData;
  // Iteratively performs equation based on cells and operators found
  for (let j = 0; j < operators.length; j++) {
    result = handleFormula(
      result,
      getCell(cells[j + 1]).cellData,
      operators[j]
    );
  }
  updateDisplayValue(result, cell);
  return result;
}

// Handles basic formulas
function handleFormula(cell1Data, cell2Data, operator) {
  // Basic math logic based on cells and operator passed into function
  switch (operator) {
    case "+":
      return parseFloat(cell1Data) + parseFloat(cell2Data);
    case "-":
      return parseFloat(cell1Data) - parseFloat(cell2Data);
    case "*":
      return parseFloat(cell1Data) * parseFloat(cell2Data);
    case "/":
      return parseFloat(cell1Data) / parseFloat(cell2Data);
  }
}

// Handles basic functions
function handleFunction(cellData, cell) {
  // Result variable is becomes equation in correct format
  let result = "";
  let firstCell;
  let lastCell;
  let cellIDs = [];

  // =SUM function
  if (cellData.startsWith("SUM")) {
    // Slice only characters within bracket. Assumes that the user inputs data correctly e.g. =SUM(A1:A2)
    let equation = cellData.slice(4, -1);
    // Split function by ":"
    let opIndex = equation.search(":");
    firstCell = getCell(equation.slice(0, opIndex));
    lastCell = getCell(equation.slice(opIndex + 1));

    // For loop through selected area only
    for (let i = firstCell.col; i <= lastCell.col; i++) {
      let cellID;
      for (let j = firstCell.row; j <= lastCell.row; j++) {
        // Finds selected cells based on cellID, appends cell to "result" equation
        cellID = numToLetter(i) + j.toString();
        cellIDs.push(cellID);
        result += cellID;
        // If not the last selected cell, add "+" to "result" equation
        if (!(i == lastCell.col && j == lastCell.row)) {
          result += "+";
        }
      }
    }
    // Parse new equation
    parseEquation(result, cell);
  } else if (cellData.startsWith("AVERAGE")) {
    // =AVERAGE function
    // Slice only characters within bracket. Assumes that the user inputs data correctly e.g. =AVERAGE(A1:A2)
    let equation = cellData.slice(8, -1);
    // Split function by ":"
    let opIndex = equation.search(":");
    firstCell = getCell(equation.slice(0, opIndex));
    lastCell = getCell(equation.slice(opIndex + 1));

    // For loop through selected area only
    let cellCount = 0;
    for (let i = firstCell.col; i <= lastCell.col; i++) {
      let cellID;
      for (let j = firstCell.row; j <= lastCell.row; j++) {
        // Finds selected cells based on cellID, appends cell to "result" equation
        cellID = numToLetter(i) + j.toString();
        cellIDs.push(cellID);
        result += cellID;
        cellCount++;
        // If last selected cell, add "+"
        if (!(i == lastCell.col && j == lastCell.row)) {
          result += "+";
        }
      }
    }
    // Parse new equation
    result = parseEquation(result, cell);
    // Add "/" and cellCount to "result" equation, call handleFormula()
    result = handleFormula(result, cellCount, "/");
    updateDisplayValue(result, cell);
  }
}

// Updates display value for current cell
function updateDisplayValue(displayVal, cell) {
  // Gets cellID and updates value
  cell.displayValue = displayVal;
  let cID = cell.colLet + cell.rowNum;
  document.getElementById(cID).value = displayVal;
}

// Updates any other cells referencing the selected cell
function updateReferences(cellData, refCell) {
  let result;
  // Finds matching cellID within matrix
  spreadsheet.forEach((row) => {
    row.forEach((cell) => {
      if (cell.cellData.includes(refCell.cellID)) {
        parseEquation(cell.cellData, cell);
      }
    });
  });
  return result;
}

// Makes cell bold
function bold() {
  let cID = selectedCell.colLet + selectedCell.rowNum;
  // Changes to bold if not already bold, otherwise revert back to original format
  if (!selectedCell.isBold) {
    document.getElementById(cID).style.fontWeight = "bold";
    selectedCell.isBold = true;
  } else {
    document.getElementById(cID).style.fontWeight = "normal";
    selectedCell.isBold = false;
  }
}

// Makes cell italics
function italics() {
  let cID = selectedCell.colLet + selectedCell.rowNum;
  // Changes to italics if not already italics, otherwise revert back to original format
  if (!selectedCell.isItalics) {
    document.getElementById(cID).style.fontStyle = "italic";
    selectedCell.isItalics = true;
  } else {
    document.getElementById(cID).style.fontStyle = "normal";
    selectedCell.isItalics = false;
  }
}

// Makes cell underlined
function underline() {
  let cID = selectedCell.colLet + selectedCell.rowNum;
  // Cell underlined if not already underlined, otherwise revert back to original format
  if (!selectedCell.isUnderlined) {
    document.getElementById(cID).style.textDecoration = "underline";
    selectedCell.isUnderlined = true;
  } else {
    document.getElementById(cID).style.textDecoration = "none";
    selectedCell.isUnderlined = false;
  }
}
