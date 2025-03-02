# Sheets Clone

A powerful spreadsheet application built with React and TypeScript that closely mimics Google Sheets functionality.

## Features

### Core Spreadsheet Features
- **Grid System**: Dynamic grid with resizable rows and columns
- **Cell Selection**: Single cell and range selection support
- **Cell Editing**: Direct cell editing with formula support
- **Keyboard Navigation**: Tab, Shift+Tab, and Enter key navigation
- **Drag and Fill**: Drag to fill cells with patterns or series

### Formula Engine
- **Mathematical Functions**:
  - `SUM`: Calculate total of cell range
  - `AVERAGE`: Calculate average of cell range
  - `MAX`: Find maximum value in range
  - `MIN`: Find minimum value in range
  - `COUNT`: Count numeric values in range

- **Data Quality Functions**:
  - `TRIM`: Remove leading/trailing whitespace
  - `UPPER`: Convert text to uppercase
  - `LOWER`: Convert text to lowercase
  - `REMOVE_DUPLICATES`: Remove duplicate values
  - `FIND_AND_REPLACE`: Search and replace text

### Formatting Options
- Font family selection
- Font size adjustment
- Text color customization
- Bold, italic, and strikethrough
- Text alignment (left, center, right)
- Cell styles and formatting

### Advanced Features
- **Function Tester**: Test formulas with sample data
- **Context Menu**: Right-click menu for cell operations
- **Multi-cell Operations**: Copy, paste, and fill
- **Grid Management**: Add/remove rows and columns
- **Sticky Headers**: Fixed row and column headers

## Getting Started

### Prerequisites

- **Node.js**: v16.x or higher (v18.x recommended)
- **npm**: v8.x or higher
- **Modern browser**: Chrome, Firefox, Edge, or Safari (latest versions)

### Installation

1. Clone the repository:
```bash
git clone https://github.com/SuprHUlk/spreadsheet-2.git
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

## Usage Guide

### Basic Operations
- Click a cell to select it
- Double-click or press any key to edit
- Drag the cell border to select multiple cells
- Use the fill handle (bottom-right corner) to auto-fill cells

### Formula Entry
1. Select a cell
2. Type '=' to start a formula
3. Enter function name (e.g., =SUM)
4. Specify cell range (e.g., =SUM(A1:A5))

### Keyboard Shortcuts
- **Tab**: Move to next cell
- **Shift + Tab**: Move to previous cell
- **Enter**: Move down one cell
- **Arrow Keys**: Navigate cells
- **F2 or Double-click**: Edit cell

### Cell Formatting
1. Select cell(s)
2. Use toolbar buttons for:
   - Font style
   - Text size
   - Color
   - Alignment
   - Bold/Italic/Strikethrough

## Tech Stack

- **React**: Frontend framework
- **TypeScript**: Type safety and better development experience
- **CSS**: Custom styling with modern CSS features

## Project Structure

```
src/
├── components/
│   ├── Cell.tsx
│   ├── ContextMenu.tsx
│   ├── FormulaBar.tsx
│   ├── FunctionTester.tsx
│   └── Toolbar.tsx
├── types.ts
├── SheetApp.tsx
└── SheetApp.css
```

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

