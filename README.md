# Spreadsheet Web Application

A web-based spreadsheet application inspired by Google Sheets. It supports mathematical functions, data validation, and various cell formatting options.

## Features

### 1. Spreadsheet Interface
- Google Sheets-like UI with a toolbar, formula bar, and cell structure.
- Drag functionality for copying cell content and formulas.
- Cell dependencies are maintained for accurate formula calculations.
- Supports basic cell formatting (bold, italics, font size, color).
- Ability to add, delete, and resize rows and columns.

### 2. Mathematical Functions
- `SUM(range)`: Calculates the sum of a range of cells.
- `AVERAGE(range)`: Returns the average of a range of cells.
- `MAX(range)`: Returns the maximum value from a range of cells.
- `MIN(range)`: Returns the minimum value from a range of cells.
- `COUNT(range)`: Counts the number of numeric cells in a given range.

### 3. Data Quality Functions
- `TRIM(cell)`: Removes leading and trailing whitespace.
- `UPPER(cell)`: Converts text to uppercase.
- `LOWER(cell)`: Converts text to lowercase.
- `REMOVE_DUPLICATES(range)`: Removes duplicate rows in a selected range.
- `FIND_AND_REPLACE(range, find, replace)`: Finds and replaces specific text.

### 4. Data Entry & Validation
- Accepts different data types: numbers, text, and dates.
- Basic validation ensures numeric cells contain only numbers.

### 5. Testing
- Users can test mathematical functions and view their results.
- Formula execution results are displayed clearly.

---

## Installation

### 1. Clone the Repository:
```sh
git clone https://github.com/your-username/spreadsheet-app.git
cd spreadsheet-app
Technologies Used
HTML, CSS, JavaScript – Frontend development.
Vite – Project bundling and development server.
Font Awesome – Icons for UI elements.
Usage
Open the application in the browser.
Click on any cell to enter data or a formula.
Use the toolbar for formatting and functions.
Test formulas using the formula bar.
