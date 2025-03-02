// Spreadsheet dimensions
const ROWS = 100;
const COLS = 26; // A to Z

// Data structure to store cell data
let spreadsheetData = {};
let activeCell = { row: 0, col: 0 };
let selectedRange = null;

// Initialize the spreadsheet
document.addEventListener('DOMContentLoaded', function() {
  createSpreadsheet();
  setupEventListeners();
});

// Create the spreadsheet structure
function createSpreadsheet() {
  createColumnHeaders();
  createRowHeaders();
  createCells();
}

// Create column headers (A, B, C, ...)
function createColumnHeaders() {
  const columnHeaders = document.getElementById('column-headers');
  
  for (let i = 0; i < COLS; i++) {
    const columnHeader = document.createElement('div');
    columnHeader.className = 'column-header';
    columnHeader.textContent = String.fromCharCode(65 + i); // A, B, C, ...
    columnHeaders.appendChild(columnHeader);
  }
}

// Create row headers (1, 2, 3, ...)
function createRowHeaders() {
  const rowHeaders = document.getElementById('row-headers');
  
  for (let i = 0; i < ROWS; i++) {
    const rowHeader = document.createElement('div');
    rowHeader.className = 'row-header';
    rowHeader.textContent = i + 1;
    rowHeaders.appendChild(rowHeader);
  }
}

// Create cells
function createCells() {
  const cellsContainer = document.getElementById('cells-container');
  cellsContainer.style.gridTemplateColumns = `repeat(${COLS}, 100px)`;
  cellsContainer.style.gridTemplateRows = `repeat(${ROWS}, 25px)`;
  
  for (let row = 0; row < ROWS; row++) {
    for (let col = 0; col < COLS; col++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      cell.dataset.row = row;
      cell.dataset.col = col;
      
      const cellDisplay = document.createElement('div');
      cellDisplay.className = 'cell-display';
      cell.appendChild(cellDisplay);
      
      const cellInput = document.createElement('input');
      cellInput.className = 'cell-input';
      cellInput.type = 'text';
      cell.appendChild(cellInput);
      
      cellsContainer.appendChild(cell);
    }
  }
}

// Set up event listeners
function setupEventListeners() {
  // Cell selection
  document.getElementById('cells-container').addEventListener('click', handleCellClick);
  
  // Formula input
  document.getElementById('formula-input').addEventListener('keydown', handleFormulaInput);
  
  // Cell input
  document.querySelectorAll('.cell-input').forEach(input => {
    input.addEventListener('blur', handleCellInputBlur);
    input.addEventListener('keydown', handleCellInputKeydown);
  });
  
  // Formula help
  document.querySelector('.formula-fx').addEventListener('click', () => {
    document.getElementById('formula-help-modal').style.display = 'block';
  });
  
  // Close modal
  document.querySelector('.close').addEventListener('click', () => {
    document.getElementById('formula-help-modal').style.display = 'none';
  });
  
  // Keyboard navigation
  document.addEventListener('keydown', handleKeyboardNavigation);
  
  // Close modal when clicking outside
  window.addEventListener('click', (event) => {
    const modal = document.getElementById('formula-help-modal');
    if (event.target === modal) {
      modal.style.display = 'none';
    }
  });
}

// Handle cell click
function handleCellClick(event) {
  const cell = event.target.closest('.cell');
  if (!cell) return;
  
  // Deselect previously selected cell
  const previouslySelected = document.querySelector('.cell.selected');
  if (previouslySelected) {
    previouslySelected.classList.remove('selected');
    previouslySelected.querySelector('.cell-input').classList.remove('active');
  }
  
  // Select new cell
  cell.classList.add('selected');
  
  // Update active cell
  activeCell.row = parseInt(cell.dataset.row);
  activeCell.col = parseInt(cell.dataset.col);
  
  // Update cell address display
  updateCellAddressDisplay();
  
  // Update formula input
  const cellId = getCellId(activeCell.row, activeCell.col);
  const cellData = spreadsheetData[cellId] || {};
  document.getElementById('formula-input').value = cellData.formula || '';
  
  // If double click, activate input
  if (event.detail === 2) {
    const input = cell.querySelector('.cell-input');
    input.classList.add('active');
    input.value = cellData.formula || cellData.displayValue || '';
    input.focus();
  }
}

// Handle formula input
function handleFormulaInput(event) {
  if (event.key === 'Enter') {
    const formula = event.target.value;
    const cellId = getCellId(activeCell.row, activeCell.col);
    
    if (formula.startsWith('=')) {
      // Process formula
      try {
        const result = evaluateFormula(formula);
        updateCellValue(cellId, result, formula);
      } catch (error) {
        updateCellValue(cellId, '#ERROR', formula);
        console.error('Formula error:', error);
      }
    } else {
      // Treat as plain text
      updateCellValue(cellId, formula, '');
    }
    
    // Update cell display
    updateCellDisplay(activeCell.row, activeCell.col);
    
    // Move to next cell
    moveSelection(1, 0);
  }
}

// Handle cell input blur
function handleCellInputBlur(event) {
  const input = event.target;
  const cell = input.closest('.cell');
  const row = parseInt(cell.dataset.row);
  const col = parseInt(cell.dataset.col);
  const cellId = getCellId(row, col);
  
  // Update cell value
  const value = input.value;
  if (value.startsWith('=')) {
    try {
      const result = evaluateFormula(value);
      updateCellValue(cellId, result, value);
    } catch (error) {
      updateCellValue(cellId, '#ERROR', value);
    }
  } else {
    updateCellValue(cellId, value, '');
  }
  
  // Update display
  updateCellDisplay(row, col);
  
  // Hide input
  input.classList.remove('active');
}

// Handle cell input keydown
function handleCellInputKeydown(event) {
  if (event.key === 'Enter') {
    event.target.blur();
    moveSelection(1, 0);
  } else if (event.key === 'Tab') {
    event.preventDefault();
    event.target.blur();
    moveSelection(0, 1);
  } else if (event.key === 'Escape') {
    event.target.blur();
  }
}

// Handle keyboard navigation
function handleKeyboardNavigation(event) {
  // Only handle if no input is active
  const activeInput = document.querySelector('.cell-input.active');
  if (activeInput) return;
  
  switch (event.key) {
    case 'ArrowUp':
      event.preventDefault();
      moveSelection(-1, 0);
      break;
    case 'ArrowDown':
      event.preventDefault();
      moveSelection(1, 0);
      break;
    case 'ArrowLeft':
      event.preventDefault();
      moveSelection(0, -1);
      break;
    case 'ArrowRight':
      event.preventDefault();
      moveSelection(0, 1);
      break;
    case 'Enter':
      event.preventDefault();
      activateCellInput();
      break;
    case 'Tab':
      event.preventDefault();
      moveSelection(0, 1);
      break;
  }
}

// Move selection
function moveSelection(rowDelta, colDelta) {
  // Calculate new position
  const newRow = Math.max(0, Math.min(ROWS - 1, activeCell.row + rowDelta));
  const newCol = Math.max(0, Math.min(COLS - 1, activeCell.col + colDelta));
  
  // Deselect current cell
  const currentCell = document.querySelector('.cell.selected');
  if (currentCell) {
    currentCell.classList.remove('selected');
    currentCell.querySelector('.cell-input').classList.remove('active');
  }
  
  // Select new cell
  activeCell.row = newRow;
  activeCell.col = newCol;
  
  const newCellElement = document.querySelector(`.cell[data-row="${newRow}"][data-col="${newCol}"]`);
  if (newCellElement) {
    newCellElement.classList.add('selected');
    
    // Scroll into view if needed
    newCellElement.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    
    // Update cell address display
    updateCellAddressDisplay();
    
    // Update formula input
    const cellId = getCellId(newRow, newCol);
    const cellData = spreadsheetData[cellId] || {};
    document.getElementById('formula-input').value = cellData.formula || '';
  }
}

// Activate cell input
function activateCellInput() {
  const cell = document.querySelector(`.cell[data-row="${activeCell.row}"][data-col="${activeCell.col}"]`);
  if (!cell) return;
  
  const input = cell.querySelector('.cell-input');
  const cellId = getCellId(activeCell.row, activeCell.col);
  const cellData = spreadsheetData[cellId] || {};
  
  input.classList.add('active');
  input.value = cellData.formula || cellData.displayValue || '';
  input.focus();
}

// Update cell address display
function updateCellAddressDisplay() {
  const colLetter = String.fromCharCode(65 + activeCell.col);
  const rowNumber = activeCell.row + 1;
  document.getElementById('current-cell').textContent = `${colLetter}${rowNumber}`;
}

// Get cell ID from row and column
function getCellId(row, col) {
  const colLetter = String.fromCharCode(65 + col);
  return `${colLetter}${row + 1}`;
}

// Get row and column from cell ID
function getCellPosition(cellId) {
  const colLetter = cellId.charAt(0);
  const col = colLetter.charCodeAt(0) - 65;
  const row = parseInt(cellId.substring(1)) - 1;
  return { row, col };
}

// Update cell value
function updateCellValue(cellId, value, formula) {
  spreadsheetData[cellId] = {
    displayValue: value,
    formula: formula
  };
  
  // Recalculate cells that depend on this one
  recalculateDependentCells(cellId);
}

// Update cell display
function updateCellDisplay(row, col) {
  const cellId = getCellId(row, col);
  const cell = document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
  if (!cell) return;
  
  const cellData = spreadsheetData[cellId] || {};
  const displayElement = cell.querySelector('.cell-display');
  displayElement.textContent = cellData.displayValue || '';
}

// Recalculate cells that depend on a changed cell
function recalculateDependentCells(changedCellId) {
  // For each cell in the spreadsheet
  Object.keys(spreadsheetData).forEach(cellId => {
    const cellData = spreadsheetData[cellId];
    
    // If the cell has a formula
    if (cellData.formula && cellData.formula.includes(changedCellId)) {
      try {
        // Re-evaluate the formula
        const result = evaluateFormula(cellData.formula);
        cellData.displayValue = result;
        
        // Update the display
        const position = getCellPosition(cellId);
        updateCellDisplay(position.row, position.col);
      } catch (error) {
        cellData.displayValue = '#ERROR';
        const position = getCellPosition(cellId);
        updateCellDisplay(position.row, position.col);
      }
    }
  });
}

// Evaluate a formula
function evaluateFormula(formula) {
  if (!formula.startsWith('=')) {
    return formula;
  }
  
  // Remove the equals sign
  const expression = formula.substring(1);
  
  // Check for built-in functions
  if (expression.toUpperCase().startsWith('SUM(')) {
    return calculateSum(expression);
  } else if (expression.toUpperCase().startsWith('AVERAGE(')) {
    return calculateAverage(expression);
  } else if (expression.toUpperCase().startsWith('MAX(')) {
    return calculateMax(expression);
  } else if (expression.toUpperCase().startsWith('MIN(')) {
    return calculateMin(expression);
  } else if (expression.toUpperCase().startsWith('COUNT(')) {
    return calculateCount(expression);
  } else if (expression.toUpperCase().startsWith('TRIM(')) {
    return applyTrim(expression);
  } else if (expression.toUpperCase().startsWith('UPPER(')) {
    return applyUpper(expression);
  } else if (expression.toUpperCase().startsWith('LOWER(')) {
    return applyLower(expression);
  } else if (expression.toUpperCase().startsWith('FIND_AND_REPLACE(')) {
    return applyFindAndReplace(expression);
  } else if (expression.toUpperCase().startsWith('REMOVE_DUPLICATES(')) {
    return applyRemoveDuplicates(expression);
  }
  
  // For simple arithmetic expressions
  try {
    // Replace cell references with their values
    const expressionWithValues = replaceReferences(expression);
    
    // Use Function constructor to evaluate the expression
    // This is safer than eval() but still has security implications
    // In a production environment, use a proper formula parser
    const result = new Function(`return ${expressionWithValues}`)();
    
    // Format the result
    return formatResult(result);
  } catch (error) {
    console.error('Error evaluating formula:', error);
    return '#ERROR';
  }
}

// Replace cell references with their values
function replaceReferences(expression) {
  // Regular expression to match cell references (e.g., A1, B2, etc.)
  const cellRefRegex = /[A-Z][0-9]+/g;
  
  return expression.replace(cellRefRegex, (match) => {
    const cellData = spreadsheetData[match] || {};
    const value = cellData.displayValue || 0;
    
    // If the value is numeric, return it directly
    if (!isNaN(value)) {
      return value;
    }
    
    // Otherwise, wrap it in quotes
    return `"${value}"`;
  });
}

// Format the result
function formatResult(result) {
  if (typeof result === 'number') {
    // Format number to a reasonable precision
    return parseFloat(result.toFixed(10)).toString();
  }
  return result;
}

// Parse a range string (e.g., "A1:B5") into an array of cell IDs
function parseRange(rangeStr) {
  // Extract the range part from the function call
  const rangeMatch = rangeStr.match(/\(([^)]+)\)/);
  if (!rangeMatch) return [];
  
  const range = rangeMatch[1];
  
  // Split the range into start and end cells
  const [startCell, endCell] = range.split(':');
  
  if (!endCell) {
    // If it's a single cell, return it
    return [startCell];
  }
  
  // Parse start and end positions
  const start = getCellPosition(startCell);
  const end = getCellPosition(endCell);
  
  // Generate all cell IDs in the range
  const cells = [];
  for (let row = start.row; row <= end.row; row++) {
    for (let col = start.col; col <= end.col; col++) {
      cells.push(getCellId(row, col));
    }
  }
  
  return cells;
}

// Get numeric values from a range of cells
function getNumericValuesFromRange(rangeStr) {
  const cells = parseRange(rangeStr);
  
  return cells.map(cellId => {
    const cellData = spreadsheetData[cellId] || {};
    const value = cellData.displayValue || '';
    
    // Convert to number or use 0 if not a number
    const numValue = parseFloat(value);
    return isNaN(numValue) ? 0 : numValue;
  });
}

// Calculate SUM
function calculateSum(expression) {
  const values = getNumericValuesFromRange(expression);
  return values.reduce((sum, value) => sum + value, 0);
}

// Calculate AVERAGE
function calculateAverage(expression) {
  const values = getNumericValuesFromRange(expression);
  if (values.length === 0) return 0;
  
  const sum = values.reduce((sum, value) => sum + value, 0);
  return sum / values.length;
}

// Calculate MAX
function calculateMax(expression) {
  const values = getNumericValuesFromRange(expression);
  if (values.length === 0) return 0;
  
  return Math.max(...values);
}

// Calculate MIN
function calculateMin(expression) {
  const values = getNumericValuesFromRange(expression);
  if (values.length === 0) return 0;
  
  return Math.min(...values);
}

// Calculate COUNT
function calculateCount(expression) {
  const cells = parseRange(expression);
  
  let count = 0;
  cells.forEach(cellId => {
    const cellData = spreadsheetData[cellId] || {};
    const value = cellData.displayValue || '';
    
    // Count only if it's a number
    if (!isNaN(parseFloat(value))) {
      count++;
    }
  });
  
  return count;
}

// Apply TRIM
function applyTrim(expression) {
  // Extract the cell reference
  const cellMatch = expression.match(/\(([^)]+)\)/);
  if (!cellMatch) return '#ERROR';
  
  const cellId = cellMatch[1];
  const cellData = spreadsheetData[cellId] || {};
  const value = cellData.displayValue || '';
  
  return value.trim();
}

// Apply UPPER
function applyUpper(expression) {
  // Extract the cell reference
  const cellMatch = expression.match(/\(([^)]+)\)/);
  if (!cellMatch) return '#ERROR';
  
  const cellId = cellMatch[1];
  const cellData = spreadsheetData[cellId] || {};
  const value = cellData.displayValue || '';
  
  return value.toUpperCase();
}

// Apply LOWER
function applyLower(expression) {
  // Extract the cell reference
  const cellMatch = expression.match(/\(([^)]+)\)/);
  if (!cellMatch) return '#ERROR';
  
  const cellId = cellMatch[1];
  const cellData = spreadsheetData[cellId] || {};
  const value = cellData.displayValue || '';
  
  return value.toLowerCase();
}

// Apply FIND_AND_REPLACE
function applyFindAndReplace(expression) {
  // Extract parameters: range, find, replace
  const paramsMatch = expression.match(/\(([^,]+),\s*"([^"]+)",\s*"([^"]+)"\)/);
  if (!paramsMatch) return '#ERROR';
  
  const [, rangeStr, findText, replaceText] = paramsMatch;
  
  // Get cells in the range
  const cells = parseRange(rangeStr);
  
  // Apply find and replace to each cell
  cells.forEach(cellId => {
    const cellData = spreadsheetData[cellId] || {};
    const value = cellData.displayValue || '';
    
    // Replace all occurrences
    const newValue = value.replace(new RegExp(findText, 'g'), replaceText);
    
    // Update the cell
    updateCellValue(cellId, newValue, '');
    
    // Update display
    const position = getCellPosition(cellId);
    updateCellDisplay(position.row, position.col);
  });
  
  return `Replaced all "${findText}" with "${replaceText}"`;
}

// Apply REMOVE_DUPLICATES
function applyRemoveDuplicates(expression) {
  // Extract the range
  const rangeMatch = expression.match(/\(([^)]+)\)/);
  if (!rangeMatch) return '#ERROR';
  
  const rangeStr = rangeMatch[1];
  
  // Get cells in the range
  const cells = parseRange(rangeStr);
  
  // Get unique values
  const values = cells.map(cellId => {
    const cellData = spreadsheetData[cellId] || {};
    return cellData.displayValue || '';
  });
  
  const uniqueValues = [...new Set(values)];
  
  // Update cells with unique values
  uniqueValues.forEach((value, index) => {
    if (index < cells.length) {
      updateCellValue(cells[index], value, '');
      
      const position = getCellPosition(cells[index]);
      updateCellDisplay(position.row, position.col);
    }
  });
  
  // Clear remaining cells
  for (let i = uniqueValues.length; i < cells.length; i++) {
    updateCellValue(cells[i], '', '');
    
    const position = getCellPosition(cells[i]);
    updateCellDisplay(position.row, position.col);
  }
  
  return `Removed ${cells.length - uniqueValues.length} duplicates`;
}

// Add some example data when the page loads
document.addEventListener('DOMContentLoaded', function() {
  // Add example data after a short delay to ensure the spreadsheet is fully created
  setTimeout(() => {
    // Add some numbers for SUM example
    updateCellValue('A1', '10', '');
    updateCellDisplay(0, 0);
    
    updateCellValue('A2', '20', '');
    updateCellDisplay(1, 0);
    
    updateCellValue('A3', '30', '');
    updateCellDisplay(2, 0);
    
    updateCellValue('A4', '40', '');
    updateCellDisplay(3, 0);
    
    updateCellValue('A5', '50', '');
    updateCellDisplay(4, 0);
    
    // Add a SUM formula
    updateCellValue('A6', '150', '=SUM(A1:A5)');
    updateCellDisplay(5, 0);
    
    // Add some text for data quality functions
    updateCellValue('C1', '  Hello World  ', '');
    updateCellDisplay(0, 2);
    
    updateCellValue('C2', 'TRIMMED TEXT', '=TRIM(C1)');
    updateCellDisplay(1, 2);
    
    updateCellValue('C3', 'lowercase text', '');
    updateCellDisplay(2, 2);
    
    updateCellValue('C4', 'LOWERCASE TEXT', '=UPPER(C3)');
    updateCellDisplay(3, 2);
  }, 500);
});