# Sheets Clone

A powerful spreadsheet application built with React that closely mimics Google Sheets functionality.

## Getting Started

### Prerequisites

- **Node.js**: v16.x or higher (v18.x recommended)
- **npm**: v8.x or higher
- **Modern browser**: Chrome, Firefox, Edge, or Safari (latest versions)

### Installation

1. Clone the repository:
```bash
git clone https://github.com/SuprHUlk/spreadsheet.git
cd spreadsheet
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm start
```
Open [http://localhost:3000](http://localhost:3000) to view the application.

4. Build for production:
```bash
npm run build
```

## Usage

- **Cell Navigation**: Use Tab/Shift+Tab to move between cells
- **Editing**: Double-click a cell or press any key when a cell is selected to edit
- **Formulas**: Start with "=" followed by a function name (e.g., =SUM(A1:A5))
- **Formatting**: Use the toolbar to change font, size, color, and alignment
- **Testing Functions**: Click the "Test" button to try out functions with sample data

## Tech Stack

### Frontend Framework
- **React**: Chosen for its efficient virtual DOM rendering and component-based architecture, which is crucial for handling large spreadsheet data with minimal performance impact.
- **TailwindCSS**: Provides utility-first CSS for rapid UI development and consistent styling.

### Data Structures

1. **Cell Data Management**
```javascript
{
  "A1": "value",
  "B2": "=SUM(A1:A5)"
}
```
- Uses an object-based structure for O(1) cell access
- Keys are cell references (e.g., "A1")
- Values store raw input including formulas

2. **Cell Styles**
```javascript
{
  "A1": {
    fontWeight: "bold",
    fontSize: "14px",
    color: "#000000"
  }
}
```
- Separate object for style management
- Efficient updates without affecting cell data

3. **Selection Management**
```javascript
{
  startRow: 0,
  endRow: 2,
  startCol: 0,
  endCol: 1
}
```
- Tracks current selection range
- Enables efficient range operations

### Key Features

1. **Formula Engine**
- Supports mathematical functions: SUM, AVERAGE, MAX, MIN, COUNT
- Data quality functions: TRIM, UPPER, LOWER, REMOVE_DUPLICATES, FIND_AND_REPLACE
- Cell reference parsing and evaluation
- Circular dependency detection

2. **UI Components**
- Toolbar with formatting options
- Formula bar with cell reference
- Context menu for row/column operations
- Function tester for validation
- Excel-like grid with resizable columns/rows

3. **Data Import/Export**
- Excel file import/export support
- Preserves formatting and formulas
- Handles multiple sheets

### Performance Optimizations

1. **Virtual Rendering**
- Only renders visible cells
- Efficient scroll handling
- Reduced memory usage

2. **Memoization**
- Caches formula results
- Prevents unnecessary recalculations
- Uses React.memo for complex components

3. **Batch Updates**
- Groups cell updates
- Minimizes re-renders
- Efficient state management

### Security Measures

1. **Formula Validation**
- Sanitizes formula inputs
- Prevents code injection
- Validates cell references

2. **Data Sanitization**
- Cleanses imported data
- Validates file types
- Handles malformed input

3. **Error Handling**
- Graceful error recovery
- User-friendly error messages
- Logging and monitoring

### Non-functional Requirements

1. **Accessibility**
- ARIA labels
- Keyboard navigation
- Screen reader support

2. **Responsive Design**
- Mobile-friendly layout
- Touch support
- Adaptive UI

3. **Browser Compatibility**
- Works across modern browsers
- Fallback support
- Progressive enhancement

