// DOM Element References
const colHeadContainer = document.getElementById("head-row-container");
const form = document.getElementById("form");
const dataContainer = document.getElementById("data-container");
const srNoContainer = document.getElementById("sr-no-container");
const disEle = document.getElementById("dis-ele");
const funExp = document.getElementById("funExp");
const downloadSymbol = document.getElementById("download-symbol");
const upload = document.getElementById("upload");
const createSheet = document.getElementById("create-sheet");
const sheets = document.getElementById("sheets");
const documentTitle = document.getElementById("document-title");

// Constants and Variables
const TOTAL_COLUMNS = 26;
const TOTAL_ROWS = 1000;
let selectedCell = null;
let activeSheetId = "sheet1";
let sheetCount = 1;
let sheetStates = {
  sheet1: {
    name: "Sheet 1",
    cells: {},
    activeCell: null
  }
};

// Default styling for cells
const DEFAULT_CELL_STYLE = {
  fontFamily: "Arial,sans-serif",
  fontSize: "16px",
  isBold: false,
  isItalic: false,
  isUnderline: false,
  textColor: "#000000",
  backgroundColor: "#ffffff",
  horizontalAlign: "left",
  innerText: ""
};

// Initialize the spreadsheet on page load
document.addEventListener('DOMContentLoaded', initializeSpreadsheet);

function initializeSpreadsheet() {
  createColumnHeaders();
  createRowNumbers();
  createCells();
  createInitialSheet();
  setupEventListeners();
}

// Create column headers (A-Z)
function createColumnHeaders() {
  for (let i = 1; i <= TOTAL_COLUMNS; i++) {
    const cell = document.createElement("div");
    cell.innerText = String.fromCharCode(i + 64);
    cell.className = "col-head";
    colHeadContainer.appendChild(cell);
  }
}

// Create row numbers (1-1000)
function createRowNumbers() {
  for (let i = 1; i <= TOTAL_ROWS; i++) {
    const cell = document.createElement("div");
    cell.className = "sr-no";
    cell.innerText = i;
    srNoContainer.appendChild(cell);
  }
}

// Create the grid cells
function createCells() {
  for (let i = 1; i <= TOTAL_ROWS; i++) {
    for (let j = 1; j <= TOTAL_COLUMNS; j++) {
      const cell = document.createElement("div");
      cell.className = "data";
      cell.contentEditable = "true";
      cell.id = `${String.fromCharCode(64 + j)}${i}`;
      cell.addEventListener("input", handleCellInput);
      cell.addEventListener("focus", handleCellFocus);
      dataContainer.appendChild(cell);
    }
  }
}

// Create the initial sheet
function createInitialSheet() {
  const sheet1 = document.createElement("div");
  sheet1.className = "sheet-tab active";
  sheet1.id = "sheet1-tab";
  sheet1.innerText = "Sheet 1";
  sheet1.dataset.sheetId = "sheet1";
  sheet1.addEventListener("click", handleSheetChange);
  sheets.appendChild(sheet1);
}

// Set up event listeners
function setupEventListeners() {
  // Form change event
  form.addEventListener("change", handleFormChange);
  
  // Formula input event
  funExp.addEventListener("keyup", handleFormulaInput);
  
  // Download event
  downloadSymbol.addEventListener("click", handleDownload);
  
  // Upload event
  upload.addEventListener("change", handleUpload);
  
  // Create new sheet event
  createSheet.addEventListener("click", handleCreateSheet);
  
  // Document title click event for renaming
  documentTitle.addEventListener("click", handleTitleClick);
  
  // Double click on cell for formula editing
  dataContainer.addEventListener("dblclick", handleCellDoubleClick);
}

// Event Handlers
function handleCellInput(e) {
  if (!selectedCell) return;
  
  const currentSheet = sheetStates[activeSheetId];
  const cellId = selectedCell.id;
  
  // Create cell state if it doesn't exist
  if (!currentSheet.cells[cellId]) {
    currentSheet.cells[cellId] = { ...DEFAULT_CELL_STYLE };
  }
  
  // Update the innerText in the state
  currentSheet.cells[cellId].innerText = selectedCell.innerText;
}

function handleCellFocus(e) {
  // Remove active class from previously selected cell
  if (selectedCell) {
    selectedCell.classList.remove("active-cell");
  }

  // Set newly selected cell
  selectedCell = e.target;
  selectedCell.classList.add("active-cell", "cell-select-animation");
  
  // Update cell reference display
  disEle.innerText = selectedCell.id;
  
  // Update current sheet's active cell
  sheetStates[activeSheetId].activeCell = selectedCell.id;
  
  // Fill form with cell's styles
  fillFormWithCellStyles();
  
  // Show formula if cell has one
  const cellState = sheetStates[activeSheetId].cells[selectedCell.id];
  if (cellState && cellState.formula) {
    funExp.value = cellState.formula;
  } else {
    funExp.value = "";
  }
}

function handleFormChange() {
  if (!selectedCell) {
    alert("Please select a cell first");
    return;
  }

  const formData = {
    fontFamily: form["font-family"].value,
    fontSize: form["font-size"].value + "px",
    isBold: form["isBold"].checked,
    isItalic: form["isItalic"].checked,
    isUnderline: form["isUnderline"].checked,
    textColor: form["text-color"].value,
    backgroundColor: form["bg-color"].value,
    horizontalAlign: form["horizontal-align"].value
  };

  // Update cell state
  const cellId = selectedCell.id;
  const currentSheet = sheetStates[activeSheetId];
  
  if (!currentSheet.cells[cellId]) {
    currentSheet.cells[cellId] = { ...DEFAULT_CELL_STYLE };
  }
  
  // Merge new styles with existing state
  currentSheet.cells[cellId] = {
    ...currentSheet.cells[cellId],
    ...formData
  };
  
  // Apply styles to the cell
  applyCellStyles(selectedCell, currentSheet.cells[cellId]);
}

function handleFormulaInput(e) {
  if (e.key !== "Enter") return;
  
  if (!selectedCell) {
    alert("Please select a cell first");
    funExp.value = "";
    return;
  }
  
  const formula = funExp.value.trim();
  
  if (!formula) return;
  
  try {
    // Store the formula in the cell state
    const cellId = selectedCell.id;
    const currentSheet = sheetStates[activeSheetId];
    
    if (!currentSheet.cells[cellId]) {
      currentSheet.cells[cellId] = { ...DEFAULT_CELL_STYLE };
    }
    
    currentSheet.cells[cellId].formula = formula;
    
    // Calculate the result
    const result = evaluateFormula(formula);
    selectedCell.innerText = result;
    currentSheet.cells[cellId].innerText = result;
    
  } catch (error) {
    alert(`Error in formula: ${error.message}`);
  }
}

function handleDownload() {
  // Create a data object with sheet states and metadata
  const data = {
    title: documentTitle.innerText,
    activeSheet: activeSheetId,
    sheets: sheetStates,
    lastModified: new Date().toISOString()
  };
  
  // Convert to JSON
  const jsonData = JSON.stringify(data, null, 2);
  
  // Create blob and trigger download
  const blob = new Blob([jsonData], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  
  const link = document.createElement("a");
  link.href = url;
  link.download = `${documentTitle.innerText.replace(/\s+/g, '_')}.json`;
  link.click();
  
  URL.revokeObjectURL(url);
}

function handleUpload(e) {
  const file = e.target.files[0];
  
  if (!file || file.type !== "application/json") {
    alert("Please upload a valid JSON file");
    return;
  }
  
  const reader = new FileReader();
  
  reader.onload = function(event) {
    try {
      const data = JSON.parse(event.target.result);
      
      // Set document title
      if (data.title) {
        documentTitle.innerText = data.title;
      }
      
      // Load sheet states
      if (data.sheets) {
        sheetStates = data.sheets;
        
        // Clear existing sheets
        sheets.innerHTML = '';
        
        // Create sheet tabs
        for (const sheetId in sheetStates) {
          createSheetTab(sheetId, sheetStates[sheetId].name);
        }
        
        // Set active sheet
        activeSheetId = data.activeSheet || Object.keys(sheetStates)[0];
        const activeSheetTab = document.querySelector(`[data-sheet-id="${activeSheetId}"]`);
        if (activeSheetTab) {
          activeSheetTab.classList.add('active');
        }
        
        // Update sheetCount
        sheetCount = Object.keys(sheetStates).length;
        
        // Load active sheet data
        loadSheetData(activeSheetId);
      }
      
    } catch (error) {
      alert(`Error loading file: ${error.message}`);
      console.error(error);
    }
  };
  
  reader.readAsText(file);
}

function handleCreateSheet() {
  sheetCount++;
  const sheetId = `sheet${sheetCount}`;
  const sheetName = `Sheet ${sheetCount}`;
  
  // Create sheet state
  sheetStates[sheetId] = {
    name: sheetName,
    cells: {},
    activeCell: null
  };
  
  // Create sheet tab
  createSheetTab(sheetId, sheetName);
  
  // Switch to new sheet
  switchToSheet(sheetId);
}

function handleSheetChange(e) {
  const sheetId = e.target.dataset.sheetId;
  switchToSheet(sheetId);
}

function handleTitleClick() {
  const newTitle = prompt("Enter a new title for your spreadsheet:", documentTitle.innerText);
  if (newTitle !== null && newTitle.trim() !== "") {
    documentTitle.innerText = newTitle.trim();
  }
}

function handleCellDoubleClick(e) {
  if (!e.target.classList.contains("data")) return;
  
  selectedCell = e.target;
  const cellState = sheetStates[activeSheetId].cells[selectedCell.id];
  
  if (cellState && cellState.formula) {
    funExp.value = cellState.formula;
    funExp.focus();
  }
}

// Helper Functions
function fillFormWithCellStyles() {
  const cellId = selectedCell.id;
  const currentSheet = sheetStates[activeSheetId];
  
  // If cell has stored state, use it; otherwise use defaults
  const cellStyle = currentSheet.cells[cellId] || DEFAULT_CELL_STYLE;
  
  // Update form fields
  form["font-family"].value = cellStyle.fontFamily;
  form["font-size"].value = parseInt(cellStyle.fontSize) || 16;
  form["isBold"].checked = cellStyle.isBold;
  form["isItalic"].checked = cellStyle.isItalic;
  form["isUnderline"].checked = cellStyle.isUnderline;
  form["text-color"].value = cellStyle.textColor;
  form["bg-color"].value = cellStyle.backgroundColor;
  
  // Set the horizontal alignment radio button
  const alignRadio = document.querySelector(`input[name="horizontal-align"][value="${cellStyle.horizontalAlign}"]`);
  if (alignRadio) {
    alignRadio.checked = true;
  }
}

function applyCellStyles(cell, styleObj) {
  if (!cell || !styleObj) return;
  
  cell.style.fontFamily = styleObj.fontFamily;
  cell.style.fontSize = styleObj.fontSize;
  cell.style.fontWeight = styleObj.isBold ? "bold" : "normal";
  cell.style.fontStyle = styleObj.isItalic ? "italic" : "normal";
  cell.style.textDecoration = styleObj.isUnderline ? "underline" : "none";
  cell.style.color = styleObj.textColor;
  cell.style.backgroundColor = styleObj.backgroundColor;
  cell.style.textAlign = styleObj.horizontalAlign;
  
  // Set content if provided
  if (styleObj.innerText !== undefined) {
    cell.innerText = styleObj.innerText;
  }
}

function evaluateFormula(formula) {
  // Simple formula evaluation for basic arithmetic
  if (formula.startsWith('=')) {
    formula = formula.substring(1);
  }
  
  // Support for cell references (e.g., =A1+B1)
  formula = convertCellReferences(formula);
  
  // Safely evaluate the expression
  try {
    // Use Function instead of eval for better security
    return new Function('return ' + formula)();
  } catch (error) {
    throw new Error(`Invalid formula: ${error.message}`);
  }
}

function convertCellReferences(formula) {
  // Replace cell references with their values
  const cellRefPattern = /[A-Z][0-9]+/g;
  return formula.replace(cellRefPattern, (match) => {
    const cell = document.getElementById(match);
    if (!cell) return '0';
    
    const cellValue = cell.innerText.trim();
    if (!cellValue || isNaN(cellValue)) return '0';
    return cellValue;
  });
}

function createSheetTab(sheetId, sheetName) {
  const sheetTab = document.createElement("div");
  sheetTab.className = "sheet-tab";
  sheetTab.dataset.sheetId = sheetId;
  sheetTab.innerText = sheetName;
  sheetTab.addEventListener("click", handleSheetChange);
  
  // Add context menu for rename/delete
  sheetTab.addEventListener("contextmenu", function(e) {
    e.preventDefault();
    showSheetContextMenu(e, sheetId);
  });
  
  sheets.appendChild(sheetTab);
  return sheetTab;
}

function switchToSheet(sheetId) {
  // Save current sheet state
  if (selectedCell) {
    sheetStates[activeSheetId].activeCell = selectedCell.id;
  }
  
  // Update active sheet
  activeSheetId = sheetId;
  
  // Update sheet tabs
  document.querySelectorAll('.sheet-tab').forEach(tab => {
    tab.classList.remove('active');
  });
  
  const activeTab = document.querySelector(`[data-sheet-id="${sheetId}"]`);
  if (activeTab) {
    activeTab.classList.add('active');
  }
  
  // Clear cell selection
  if (selectedCell) {
    selectedCell.classList.remove('active-cell');
    selectedCell = null;
  }
  
  // Clear formula input
  funExp.value = "";
  
  // Load sheet data
  loadSheetData(sheetId);
}

function loadSheetData(sheetId) {
  // Clear all cells
  const cells = document.querySelectorAll('.data');
  cells.forEach(cell => {
    cell.innerText = '';
    cell.removeAttribute('style');
  });
  
  // Apply stored styles and content
  const currentSheet = sheetStates[sheetId];
  if (!currentSheet) return;
  
  for (const cellId in currentSheet.cells) {
    const cell = document.getElementById(cellId);
    if (cell) {
      applyCellStyles(cell, currentSheet.cells[cellId]);
    }
  }
  
  // Restore active cell
  if (currentSheet.activeCell) {
    const cell = document.getElementById(currentSheet.activeCell);
    if (cell) {
      cell.click();
    }
  }
}

function showSheetContextMenu(e, sheetId) {
  // Remove any existing context menu
  const oldMenu = document.getElementById('sheet-context-menu');
  if (oldMenu) {
    oldMenu.remove();
  }
  
  // Create context menu
  const menu = document.createElement('div');
  menu.id = 'sheet-context-menu';
  menu.style.position = 'absolute';
  menu.style.left = `${e.pageX}px`;
  menu.style.top = `${e.pageY}px`;
  menu.style.backgroundColor = '#fff';
  menu.style.border = '1px solid #ccc';
  menu.style.borderRadius = '4px';
  menu.style.boxShadow = '0 2px 10px rgba(0,0,0,0.2)';
  menu.style.padding = '5px 0';
  menu.style.zIndex = '1000';
  
  // Rename option
  const renameOption = document.createElement('div');
  renameOption.innerText = 'Rename';
  renameOption.style.padding = '8px 15px';
  renameOption.style.cursor = 'pointer';
  renameOption.style.transition = 'background-color 0.2s';
  
  renameOption.addEventListener('mouseover', () => {
    renameOption.style.backgroundColor = '#f1f3f4';
  });
  
  renameOption.addEventListener('mouseout', () => {
    renameOption.style.backgroundColor = 'transparent';
  });
  
  renameOption.addEventListener('click', () => {
    renameSheet(sheetId);
    menu.remove();
  });
  
  // Delete option (only if there's more than one sheet)
  if (Object.keys(sheetStates).length > 1) {
    const deleteOption = document.createElement('div');
    deleteOption.innerText = 'Delete';
    deleteOption.style.padding = '8px 15px';
    deleteOption.style.cursor = 'pointer';
    deleteOption.style.transition = 'background-color 0.2s';
    
    deleteOption.addEventListener('mouseover', () => {
      deleteOption.style.backgroundColor = '#f1f3f4';
    });
    
    deleteOption.addEventListener('mouseout', () => {
      deleteOption.style.backgroundColor = 'transparent';
    });
    
    deleteOption.addEventListener('click', () => {
      deleteSheet(sheetId);
      menu.remove();
    });
    
    menu.appendChild(deleteOption);
  }
  
  menu.appendChild(renameOption);
  document.body.appendChild(menu);
  
  // Close menu when clicking elsewhere
  document.addEventListener('click', function closeMenu(e) {
    if (e.target !== menu && !menu.contains(e.target)) {
      menu.remove();
      document.removeEventListener('click', closeMenu);
    }
  });
}

function renameSheet(sheetId) {
  const currentName = sheetStates[sheetId].name;
  const newName = prompt("Enter a new name for this sheet:", currentName);
  
  if (newName !== null && newName.trim() !== "") {
    // Update state
    sheetStates[sheetId].name = newName.trim();
    
    // Update tab text
    const sheetTab = document.querySelector(`[data-sheet-id="${sheetId}"]`);
    if (sheetTab) {
      sheetTab.innerText = newName.trim();
    }
  }
}

function deleteSheet(sheetId) {
  if (Object.keys(sheetStates).length <= 1) {
    alert("Cannot delete the only sheet");
    return;
  }
  
  if (confirm(`Are you sure you want to delete "${sheetStates[sheetId].name}"?`)) {
    // Remove sheet tab
    const sheetTab = document.querySelector(`[data-sheet-id="${sheetId}"]`);
    if (sheetTab) {
      sheetTab.remove();
    }
    
    // Remove sheet state
    delete sheetStates[sheetId];
    
    // Switch to another sheet if the active sheet was deleted
    if (activeSheetId === sheetId) {
      const newActiveSheetId = Object.keys(sheetStates)[0];
      switchToSheet(newActiveSheetId);
    }
  }
}

// Export helper functions for testing
window.spreadsheetHelpers = {
  evaluateFormula,
  convertCellReferences,
  applyCellStyles
};