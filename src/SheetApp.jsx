import React, { useState, useCallback, useRef, useEffect } from 'react';
import Toolbar from './components/Toolbar';
import FormulaBar from './components/FormulaBar';
import Cell from './components/Cell';
import FunctionTester from './components/FunctionTester';
import './SheetApp.css';
import * as XLSX from 'xlsx';

const DEFAULT_ROWS = 50;
const DEFAULT_COLS = 26;

// Helper functions for cell references and formulas
const getCellName = (row, col) => String.fromCharCode(65 + col) + (row + 1);

const parseCell = (cellStr) => {
  const col = cellStr.charCodeAt(0) - 65;
  const row = parseInt(cellStr.slice(1), 10) - 1;
  return { row, col };
};

const parseRange = (rangeStr) => {
  const parts = rangeStr.split(':');
  return parts.length === 1 ? [parts[0], parts[0]] : parts;
};

function SheetApp() {
  // State management
  const [numRows, setNumRows] = useState(DEFAULT_ROWS);
  const [numCols, setNumCols] = useState(DEFAULT_COLS);
  const [cells, setCells] = useState({});
  const [cellStyles, setCellStyles] = useState({});
  const [selectedCell, setSelectedCell] = useState('A1');
  const [selectionRange, setSelectionRange] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [dragStart, setDragStart] = useState(null);
  const gridRef = useRef(null);
  const [fontFamily, setFontFamily] = useState('Arial');
  const [dragSource, setDragSource] = useState(null);
  const [fillSource, setFillSource] = useState(null);
  const [fillTarget, setFillTarget] = useState(null);
  const [formulaText, setFormulaText] = useState(null);
  const [columnWidths, setColumnWidths] = useState({});
  const [rowHeights, setRowHeights] = useState({});
  const [isResizing, setIsResizing] = useState(false);
  const resizeStartRef = useRef(null);
  const resizeTypeRef = useRef(null);
  const resizeIndexRef = useRef(null);
  const [resizeLinePosition, setResizeLinePosition] = useState(null);
  const [showFunctionTester, setShowFunctionTester] = useState(false);

  // Cell style management
  const updateCellStyle = (cellName, styleUpdate) => {
    setCellStyles(prev => ({
      ...prev,
      [cellName]: { 
        ...(prev[cellName] || {}), 
        ...styleUpdate 
      }
    }));
  };

  // Mathematical functions
  const mathFunctions = {
    SUM: (values) => values.reduce((a, b) => a + b, 0),
    AVERAGE: (values) => values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0,
    MAX: (values) => values.length ? Math.max(...values) : 0,
    MIN: (values) => values.length ? Math.min(...values) : 0,
    COUNT: (values) => values.filter(value => {
      const num = parseFloat(value);
      return !isNaN(num) && isFinite(num);
    }).length
  };

  // Data quality functions
  const dataFunctions = {
    TRIM: (value) => {
      if (Array.isArray(value)) return value[0].trim();
      return value.trim();
    },
    UPPER: (value) => {
      if (Array.isArray(value)) return value[0].toUpperCase();
      return value.toUpperCase();
    },
    LOWER: (value) => {
      if (Array.isArray(value)) return value[0].toLowerCase();
      return value.toLowerCase();
    },
    REMOVE_DUPLICATES: (range) => {
      if (!Array.isArray(range)) return range;
      
      // Convert range to string for comparison
      const stringifyRow = (row) => {
        if (Array.isArray(row)) {
          return row.join('|');
        }
        return String(row);
      };

      // Create a Set of stringified rows to remove duplicates
      const uniqueRows = new Set(range.map(stringifyRow));
      
      // Convert back to original format
      return Array.from(uniqueRows).map(str => {
        if (str.includes('|')) {
          return str.split('|');
        }
        return str;
      });
    },
    FIND_AND_REPLACE: (text, find, replace) => {
      if (!text || !find) return text;
      
      // Handle array of values (range of cells)
      if (Array.isArray(text)) {
        return text.map(value => 
          String(value).replace(new RegExp(find, 'g'), replace || '')
        );
      }
      
      return String(text).replace(new RegExp(find, 'g'), replace || '');
    }
  };

  // Selection handling
  const handleCellSelect = (cellName, e) => {
    if (e?.shiftKey && selectedCell) {
      // Handle range selection with shift key
      const start = parseCell(selectedCell);
      const end = parseCell(cellName);
      
      const startRow = Math.min(start.row, end.row);
      const endRow = Math.max(start.row, end.row);
      const startCol = Math.min(start.col, end.col);
      const endCol = Math.max(start.col, end.col);
      
      setSelectionRange({ startRow, endRow, startCol, endCol });
    } else {
      setSelectedCell(cellName);
      setSelectionRange(null);
    }
  };

  // Mouse event handlers for drag selection
  const handleMouseDown = (cellName, e) => {
    if (!e.shiftKey) {
      setIsDragging(true);
      setDragStart(cellName);
      setSelectedCell(cellName);
      setSelectionRange(null);
    }
  };

  const handleMouseMove = (cellName) => {
    if (isDragging && dragStart) {
      const start = parseCell(dragStart);
      const end = parseCell(cellName);
      
      const startRow = Math.min(start.row, end.row);
      const endRow = Math.max(start.row, end.row);
      const startCol = Math.min(start.col, end.col);
      const endCol = Math.max(start.col, end.col);
      
      setSelectionRange({ startRow, endRow, startCol, endCol });
      setSelectedCell(cellName);
    }
  };

  const handleMouseUp = () => {
    setIsDragging(false);
    setDragStart(null);
  };

  // Effect for mouse up handler
  useEffect(() => {
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  // Function to update cell references in a formula
  const updateFormulaReferences = useCallback((formula, fromCell, toCell) => {
    if (!formula?.startsWith('=')) return formula;
    
    const expression = formula.slice(1);
    return '=' + expression.replace(/[A-Z]+[0-9]+/g, (ref) => {
      if (ref === fromCell) return toCell;
      return ref;
    });
  }, []);

  // Function to adjust formula for fill operation
  const adjustFormulaForFill = useCallback((formula, sourceCell, targetCell) => {
    if (!formula?.startsWith('=')) return formula;
    
    const source = parseCell(sourceCell);
    const target = parseCell(targetCell);
    const rowDiff = target.row - source.row;
    const colDiff = target.col - source.col;
    
    const expression = formula.slice(1);
    return '=' + expression.replace(/[A-Z]+[0-9]+/g, (ref) => {
      const refCell = parseCell(ref);
      const newRow = refCell.row + rowDiff;
      const newCol = refCell.col + colDiff;
      
      // Check if the new reference is valid
      if (newRow >= 0 && newRow < numRows && newCol >= 0 && newCol < numCols) {
        return getCellName(newRow, newCol);
      }
      return ref; // Keep original reference if new one would be invalid
    });
  }, [numRows, numCols]);

  // Formula evaluation
  const evaluateFormula = useCallback((formula, cellName) => {
    if (!formula.startsWith('=')) return formula;
    
    const expression = formula.slice(1);
    const funcMatch = expression.match(/(\w+)\((.*)\)/);
    
    if (!funcMatch) {
      // Handle basic arithmetic
      try {
        const evalExpr = expression.replace(/[A-Z]\d+/g, (ref) => {
          if (ref === cellName) return 0; // Prevent circular references
          const value = getCellDisplayValue(ref);
          return isNaN(parseFloat(value)) ? 0 : parseFloat(value);
        });
        return Function('"use strict";return (' + evalExpr + ')')();
      } catch (e) {
        return '#ERROR!';
      }
    }

    const [_, func, args] = funcMatch;
    const functionName = func.toUpperCase();

    if (mathFunctions[functionName]) {
      const [startRef, endRef] = parseRange(args);
      const start = parseCell(startRef);
      const end = parseCell(endRef);
      const values = [];

      for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
        for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
          const ref = getCellName(row, col);
          if (ref === cellName) continue; // Prevent circular references
          const value = getCellDisplayValue(ref);
          const numValue = parseFloat(value);
          if (!isNaN(numValue)) {
            values.push(numValue);
          }
        }
      }

      return mathFunctions[functionName](values);
    }

    if (dataFunctions[functionName]) {
      const argValues = args.split(',').map(arg => {
        const trimmed = arg.trim();
        if (trimmed.match(/^[A-Z]\d+$/)) {
          return getCellDisplayValue(trimmed);
        }
        return trimmed.replace(/^['"]|['"]$/g, '');
      });

      return dataFunctions[functionName](...argValues);
    }

    return '#UNKNOWN_FUNCTION';
  }, [cells]);

  const getCellDisplayValue = useCallback((cellName) => {
    const value = cells[cellName] || '';
    if (value.startsWith('=')) {
      try {
        const result = evaluateFormula(value, cellName);
        // Format numbers to avoid unnecessary decimal places
        return typeof result === 'number' ? 
          (Number.isInteger(result) ? result : result.toFixed(2)) 
          : result;
      } catch (e) {
        return '#ERROR!';
      }
    }
    return value;
  }, [cells, evaluateFormula]);

  // Cell operations
  const handleCellChange = (cellName, value) => {
    setCells(prev => ({ ...prev, [cellName]: value }));
  };

  // Row and column operations
  const handleAddRow = useCallback((index) => {
    // Update cells state to accommodate new row
    const newCells = { ...cells };
    // Shift all cells below the insertion point up by one row
    Object.entries(newCells).forEach(([cellName, value]) => {
      const { row, col } = parseCell(cellName);
      if (row > index) {
        const newCellName = getCellName(row + 1, col);
        newCells[newCellName] = value;
        delete newCells[cellName];
      }
    });
    setNumRows(numRows + 1);
    setCells(newCells);
  }, [cells, numRows]);

  const handleDeleteRow = useCallback((index) => {
    if (numRows > 1) {
      const newCells = { ...cells };
      // Remove cells in the deleted row and shift cells up
      Object.entries(newCells).forEach(([cellName, value]) => {
        const { row, col } = parseCell(cellName);
        if (row === index) {
          delete newCells[cellName];
        } else if (row > index) {
          const newCellName = getCellName(row - 1, col);
          newCells[newCellName] = value;
          delete newCells[cellName];
        }
      });
      setNumRows(numRows - 1);
      setCells(newCells);
    }
  }, [cells, numRows]);

  const handleAddColumn = useCallback((index) => {
    const newCells = { ...cells };
    // Shift all cells to the right of the insertion point by one column
    Object.entries(newCells).forEach(([cellName, value]) => {
      const { row, col } = parseCell(cellName);
      if (col > index) {
        const newCellName = getCellName(row, col + 1);
        newCells[newCellName] = value;
        delete newCells[cellName];
      }
    });
    setNumCols(numCols + 1);
    setCells(newCells);
  }, [cells, numCols]);

  const handleDeleteColumn = useCallback((index) => {
    if (numCols > 1) {
      const newCells = { ...cells };
      // Remove cells in the deleted column and shift cells left
      Object.entries(newCells).forEach(([cellName, value]) => {
        const { row, col } = parseCell(cellName);
        if (col === index) {
          delete newCells[cellName];
        } else if (col > index) {
          const newCellName = getCellName(row, col - 1);
          newCells[newCellName] = value;
          delete newCells[cellName];
        }
      });
      setNumCols(numCols - 1);
      setCells(newCells);
    }
  }, [cells, numCols]);

  // Get cell class based on selection state
  const getCellClass = (row, col) => {
    const cellName = getCellName(row, col);
    
    if (cellName === selectedCell) {
      return 'selected primary';
    }
    
    if (selectionRange) {
      const { startRow, endRow, startCol, endCol } = selectionRange;
      if (row >= startRow && row <= endRow && col >= startCol && col <= endCol) {
        return 'selected';
      }
    }
    
    return '';
  };

  // Style handlers
  const handleBold = () => {
    if (!selectedCell) return;
    updateCellStyle(selectedCell, { 
      fontWeight: cellStyles[selectedCell]?.fontWeight === 'bold' ? 'normal' : 'bold' 
    });
  };

  const handleItalic = () => {
    if (!selectedCell) return;
    updateCellStyle(selectedCell, { 
      fontStyle: cellStyles[selectedCell]?.fontStyle === 'italic' ? 'normal' : 'italic' 
    });
  };

  const handleFontSize = (size) => {
    if (!selectedCell) return;
    updateCellStyle(selectedCell, { fontSize: size });
  };

  const handleColor = (color) => {
    if (!selectedCell) return;
    updateCellStyle(selectedCell, { color });
  };

  const handleFontFamily = (font) => {
    if (!selectedCell) return;
    setFontFamily(font);
    updateCellStyle(selectedCell, { fontFamily: font });
  };

  const handleStrikethrough = () => {
    if (!selectedCell) return;
    updateCellStyle(selectedCell, { 
      textDecoration: cellStyles[selectedCell]?.textDecoration === 'line-through' ? 'none' : 'line-through' 
    });
  };

  const handleAlignment = (alignment) => {
    if (!selectedCell) return;
    
    setCellStyles(prev => ({
      ...prev,
      [selectedCell]: {
        ...prev[selectedCell],
        textAlign: alignment
      }
    }));
  };

  // Drag and drop handlers
  const handleCellDragStart = (cellName) => {
    setDragSource(cellName);
    setSelectedCell(cellName);
  };

  const handleCellDragEnd = () => {
    if (dragSource && selectedCell && dragSource !== selectedCell) {
      // Move cell content
      const sourceContent = cells[dragSource];
      const sourceStyle = cellStyles[dragSource];
      const targetContent = cells[selectedCell];
      const targetStyle = cellStyles[selectedCell];
      
      setCells(prev => {
        const newCells = { ...prev };
        // Update the moved cell content and adjust any formulas
        newCells[selectedCell] = updateFormulaReferences(sourceContent, dragSource, selectedCell);
        newCells[dragSource] = updateFormulaReferences(targetContent, selectedCell, dragSource);
        
        // Update any formulas that reference the moved cells
        Object.entries(newCells).forEach(([cellName, value]) => {
          if (value?.startsWith('=')) {
            newCells[cellName] = updateFormulaReferences(
              updateFormulaReferences(value, dragSource, selectedCell),
              selectedCell, dragSource
            );
          }
        });
        
        return newCells;
      });

      setCellStyles(prev => {
        const newStyles = { ...prev };
        newStyles[selectedCell] = sourceStyle;
        newStyles[dragSource] = targetStyle;
        return newStyles;
      });
    }
    setDragSource(null);
  };

  // Fill handle handlers
  const handleFillDragStart = (cellName) => {
    setFillSource(cellName);
    setSelectedCell(cellName);
  };

  const getCellFromPoint = (x, y) => {
    const table = gridRef.current;
    if (!table) return null;

    const cells = table.getElementsByTagName('td');
    const tableRect = table.getBoundingClientRect();
    
    for (let cell of cells) {
      const cellRect = cell.getBoundingClientRect();
      if (x >= cellRect.left && x <= cellRect.right &&
          y >= cellRect.top && y <= cellRect.bottom &&
          x >= tableRect.left && x <= tableRect.right &&
          y >= tableRect.top && y <= tableRect.bottom) {
        return cell.getAttribute('data-cell-name');
      }
    }
    return null;
  };

  const handleFillDragMove = (e) => {
    const targetCell = getCellFromPoint(e.clientX, e.clientY);
    if (targetCell && fillSource) {
      setFillTarget(targetCell);
      
      const start = parseCell(fillSource);
      const end = parseCell(targetCell);
      
      setSelectionRange({
        startRow: Math.min(start.row, end.row),
        endRow: Math.max(start.row, end.row),
        startCol: Math.min(start.col, end.col),
        endCol: Math.max(start.col, end.col)
      });
    }
  };

  const handleFillDragEnd = () => {
    if (fillSource && fillTarget) {
      const sourceValue = cells[fillSource];
      const sourceStyle = cellStyles[fillSource];
      
      if (typeof sourceValue === 'string' && sourceValue.match(/^\d+$/)) {
        // Numeric sequence
        const startNum = parseInt(sourceValue);
        const start = parseCell(fillSource);
        const end = parseCell(fillTarget);
        
        const newCells = { ...cells };
        const newStyles = { ...cellStyles };
        
        const isVertical = start.col === end.col;
        let counter = 0;
        
        if (isVertical) {
          for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
            const cellName = getCellName(row, start.col);
            if (cellName !== fillSource) {
              newCells[cellName] = String(startNum + counter);
              newStyles[cellName] = { ...sourceStyle };
              counter++;
            }
          }
        } else {
          for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
            const cellName = getCellName(start.row, col);
            if (cellName !== fillSource) {
              newCells[cellName] = String(startNum + counter);
              newStyles[cellName] = { ...sourceStyle };
              counter++;
            }
          }
        }
        
        setCells(newCells);
        setCellStyles(newStyles);
      } else if (sourceValue?.startsWith('=')) {
        // Formula fill - adjust references
        const start = parseCell(fillSource);
        const end = parseCell(fillTarget);
        
        const newCells = { ...cells };
        const newStyles = { ...cellStyles };
        
        for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
          for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
            const cellName = getCellName(row, col);
            if (cellName !== fillSource) {
              newCells[cellName] = adjustFormulaForFill(sourceValue, fillSource, cellName);
              newStyles[cellName] = { ...sourceStyle };
            }
          }
        }
        
        setCells(newCells);
        setCellStyles(newStyles);
      } else {
        // Copy value
        const start = parseCell(fillSource);
        const end = parseCell(fillTarget);
        
        const newCells = { ...cells };
        const newStyles = { ...cellStyles };
        
        for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
          for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
            const cellName = getCellName(row, col);
            if (cellName !== fillSource) {
              newCells[cellName] = sourceValue;
              newStyles[cellName] = { ...sourceStyle };
            }
          }
        }
        
        setCells(newCells);
        setCellStyles(newStyles);
      }
    }
    
    setFillSource(null);
    setFillTarget(null);
    setSelectionRange(null);
  };

  // Add this handler function inside the SheetApp component
  const handleFormulaSelect = (formula) => {
    if (!selectedCell || !formula) return;
    setFormulaText(formula);
  };

  // Add these new handlers
  const handleResizeStart = (e, type, index) => {
    e.preventDefault();
    e.stopPropagation();
    setIsResizing(true);
    
    const rect = e.currentTarget.closest('th').getBoundingClientRect();
    const tableRect = gridRef.current.getBoundingClientRect();
    
    let initialPosition;
    if (type === 'column') {
      initialPosition = e.clientX;
    } else {
      initialPosition = e.clientY;
    }
    
    resizeStartRef.current = {
      position: type === 'column' ? e.clientX : e.clientY,
      size: type === 'column' ? (columnWidths[index] || 100) : (rowHeights[index] || 24),
      initialPosition
    };
    resizeTypeRef.current = type;
    resizeIndexRef.current = index;
    setResizeLinePosition(initialPosition);
    
    document.body.classList.add('resizing');
    if (type === 'row') {
      document.body.classList.add('row');
    }
  };

  const handleResizeMove = useCallback((e) => {
    if (!isResizing || !resizeStartRef.current) return;
    
    const startData = resizeStartRef.current;
    const currentPos = resizeTypeRef.current === 'column' ? e.clientX : e.clientY;
    
    setResizeLinePosition(currentPos);
  }, [isResizing]);

  useEffect(() => {
    if (isResizing) {
      const handleMove = (e) => {
        e.preventDefault();
        e.stopPropagation();
        handleResizeMove(e);
      };
      
      const handleUp = (e) => {
        e.preventDefault();
        e.stopPropagation();
        
        const type = resizeTypeRef.current;
        const index = resizeIndexRef.current;
        const startData = resizeStartRef.current;
        const endPos = type === 'column' ? e.clientX : e.clientY;
        const diff = endPos - startData.position;
        
        if (type === 'column') {
          const newWidth = Math.max(50, startData.size + diff);
          setColumnWidths(prev => ({
            ...prev,
            [index]: newWidth
          }));
        } else {
          const newHeight = Math.max(24, startData.size + diff);
          setRowHeights(prev => ({
            ...prev,
            [index]: newHeight
          }));
        }
        
        setIsResizing(false);
        setResizeLinePosition(null);
        resizeStartRef.current = null;
        resizeTypeRef.current = null;
        resizeIndexRef.current = null;
        document.body.classList.remove('resizing', 'row');
      };
      
      window.addEventListener('mousemove', handleMove, { capture: true });
      window.addEventListener('mouseup', handleUp, { capture: true });
      
      return () => {
        window.removeEventListener('mousemove', handleMove, { capture: true });
        window.removeEventListener('mouseup', handleUp, { capture: true });
      };
    }
  }, [isResizing, handleResizeMove]);

  // Save spreadsheet data
  const handleSave = () => {
    // Create a new workbook
    const wb = XLSX.utils.book_new();
    
    // Convert cells data to array format for Excel
    const data = [];
    let maxRow = 0;
    let maxCol = 0;
    
    // Find the maximum used row and column
    Object.keys(cells).forEach(cellName => {
      const { row, col } = parseCell(cellName);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    });
    
    // Create data array with all cells (including empty ones)
    for (let row = 0; row <= maxRow; row++) {
      const rowData = [];
      for (let col = 0; col <= maxCol; col++) {
        const cellName = getCellName(row, col);
        rowData.push(getCellDisplayValue(cellName));
      }
      data.push(rowData);
    }
    
    // Create worksheet from data
    const ws = XLSX.utils.aoa_to_sheet(data);
    
    // Add column widths
    ws['!cols'] = Array(maxCol + 1).fill(null).map((_, i) => ({
      wch: columnWidths[i] ? Math.floor(columnWidths[i] / 7) : 14 // Convert pixels to Excel width units (approx)
    }));
    
    // Add row heights
    ws['!rows'] = Array(maxRow + 1).fill(null).map((_, i) => ({
      hpt: rowHeights[i] ? Math.floor(rowHeights[i] * 0.75) : 18 // Convert pixels to Excel height units (approx)
    }));
    
    // Add styles
    Object.keys(cellStyles).forEach(cellName => {
      const style = cellStyles[cellName];
      const addr = XLSX.utils.decode_cell(cellName);
      const cell = ws[XLSX.utils.encode_cell(addr)];
      if (cell) {
        cell.s = {
          font: {
            bold: style.fontWeight === 'bold',
            italic: style.fontStyle === 'italic',
            name: style.fontFamily,
            sz: parseInt(style.fontSize) || 11,
            color: { rgb: style.color?.replace('#', '') }
          }
        };
      }
    });
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    
    // Generate Excel file and trigger download
    XLSX.writeFile(wb, 'spreadsheet.xlsx');
  };

  // Load spreadsheet data
  const handleLoad = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.onchange = (e) => {
      const file = e.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (event) => {
          try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get the first worksheet
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // Convert worksheet to array of arrays
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            // Clear existing data
            setCells({});
            setCellStyles({});
            
            // Update dimensions
            setNumRows(Math.max(DEFAULT_ROWS, jsonData.length));
            setNumCols(Math.max(DEFAULT_COLS, Math.max(...jsonData.map(row => row.length))));
            
            // Convert data to our cell format
            const newCells = {};
            const newStyles = {};
            
            jsonData.forEach((row, rowIndex) => {
              row.forEach((cellValue, colIndex) => {
                if (cellValue !== null && cellValue !== undefined) {
                  const cellName = getCellName(rowIndex, colIndex);
                  newCells[cellName] = cellValue.toString();
                  
                  // Extract cell style if available
                  const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
                  if (cell?.s) {
                    const style = {};
                    if (cell.s.font) {
                      if (cell.s.font.bold) style.fontWeight = 'bold';
                      if (cell.s.font.italic) style.fontStyle = 'italic';
                      if (cell.s.font.name) style.fontFamily = cell.s.font.name;
                      if (cell.s.font.sz) style.fontSize = `${cell.s.font.sz}px`;
                      if (cell.s.font.color?.rgb) style.color = `#${cell.s.font.color.rgb}`;
                    }
                    if (Object.keys(style).length > 0) {
                      newStyles[cellName] = style;
                    }
                  }
                }
              });
            });
            
            // Update column widths
            const newColumnWidths = {};
            if (worksheet['!cols']) {
              worksheet['!cols'].forEach((col, index) => {
                if (col?.wch) {
                  newColumnWidths[index] = col.wch * 7; // Convert Excel width units to pixels (approx)
                }
              });
            }
            
            // Update row heights
            const newRowHeights = {};
            if (worksheet['!rows']) {
              worksheet['!rows'].forEach((row, index) => {
                if (row?.hpt) {
                  newRowHeights[index] = row.hpt / 0.75; // Convert Excel height units to pixels (approx)
                }
              });
            }
            
            setCells(newCells);
            setCellStyles(newStyles);
            setColumnWidths(newColumnWidths);
            setRowHeights(newRowHeights);
            
          } catch (error) {
            console.error('Error loading spreadsheet:', error);
            alert('Error loading spreadsheet: Invalid file format');
          }
        };
        reader.readAsArrayBuffer(file);
      }
    };
    input.click();
  };

  // Add this function to handle tab navigation
  const handleTabNavigation = (cellName, isShiftKey, isEnterKey = false) => {
    const { row, col } = parseCell(cellName);
    
    let nextRow = row;
    let nextCol = col;
    
    if (isEnterKey) {
      // Move down on Enter key
      nextRow = row + 1;
    } else if (isShiftKey) {
      // Move left on Shift+Tab
      if (col > 0) {
        nextCol = col - 1;
      } else if (row > 0) {
        // Move to the end of the previous row
        nextRow = row - 1;
        nextCol = numCols - 1;
      }
    } else {
      // Move right on Tab
      if (col < numCols - 1) {
        nextCol = col + 1;
      } else if (row < numRows - 1) {
        // Move to the beginning of the next row
        nextRow = row + 1;
        nextCol = 0;
      }
    }
    
    // Ensure we don't go out of bounds
    nextRow = Math.max(0, Math.min(nextRow, numRows - 1));
    nextCol = Math.max(0, Math.min(nextCol, numCols - 1));
    
    // Select the next cell
    const nextCellName = getCellName(nextRow, nextCol);
    handleCellSelect(nextCellName);
  };

  return (
    <div className="sheet-app h-screen flex flex-col">
      <Toolbar
        onBold={handleBold}
        onItalic={handleItalic}
        onFontSize={handleFontSize}
        onColor={handleColor}
        onFontFamily={handleFontFamily}
        onFormulaSelect={handleFormulaSelect}
        selectedCell={selectedCell}
        cellStyles={cellStyles[selectedCell] || {}}
        currentFont={fontFamily}
        onSave={handleSave}
        onLoad={handleLoad}
        onAddRow={handleAddRow}
        onDeleteRow={handleDeleteRow}
        onAddColumn={handleAddColumn}
        onDeleteColumn={handleDeleteColumn}
        onAlignment={handleAlignment}
        onOpenFunctionTester={() => setShowFunctionTester(true)}
      />
      
      <FormulaBar
        value={cells[selectedCell] || ''}
        onChange={(value) => handleCellChange(selectedCell, value)}
        selectedCell={selectedCell}
        formulaText={formulaText}
        onCellSelect={(cell) => {
          setSelectedCell(cell);
          setSelectionRange(null);
        }}
      />
      
      <div className="spreadsheet-container flex-1 overflow-auto relative" ref={gridRef}>
        {isResizing && resizeLinePosition && (
          <div 
            className={`resize-reference-line ${resizeTypeRef.current}`}
            style={{
              [resizeTypeRef.current === 'column' ? 'left' : 'top']: `${resizeLinePosition}px`
            }}
          />
        )}
        <table className="spreadsheet w-full table-fixed">
        <thead>
          <tr>
              <th className="w-10 bg-gray-100"></th>
              {Array.from({ length: numCols }, (_, col) => (
                <th 
                  key={col} 
                  className="bg-gray-100 px-2 py-1 text-sm relative"
                  style={{ minWidth: columnWidths[col] || 100 }}
                >
                  {String.fromCharCode(65 + col)}
                  <div 
                    className="col-resize-handle"
                    onMouseDown={(e) => handleResizeStart(e, 'column', col)}
                  />
                </th>
            ))}
          </tr>
        </thead>
        <tbody>
            {Array.from({ length: numRows }, (_, row) => (
              <tr key={row} style={{ height: rowHeights[row] || 24 }}>
                <th className="bg-gray-100 px-2 py-1 text-sm relative">
                  {row + 1}
                  <div 
                    className="row-resize-handle"
                    onMouseDown={(e) => handleResizeStart(e, 'row', row)}
                  />
                </th>
                {Array.from({ length: numCols }, (_, col) => {
                const cellName = getCellName(row, col);
                  const cellStyle = cellStyles[cellName] || {};
                return (
                    <Cell
                    key={`${row}-${col}`}
                      name={cellName}
                      value={cells[cellName] || ''}
                      displayValue={getCellDisplayValue(cellName)}
                      onChange={(value) => handleCellChange(cellName, value)}
                      onSelect={(e) => handleCellSelect(cellName, e)}
                      onMouseDown={(e) => handleMouseDown(cellName, e)}
                      onMouseMove={() => handleMouseMove(cellName)}
                      isSelected={getCellClass(row, col)}
                      styles={cellStyle}
                      onDragStart={handleCellDragStart}
                      onDragEnd={handleCellDragEnd}
                      onFillDragStart={handleFillDragStart}
                      onFillDragMove={handleFillDragMove}
                      onFillDragEnd={handleFillDragEnd}
                      onAddRow={() => handleAddRow(row)}
                      onDeleteRow={() => handleDeleteRow(row)}
                      onAddColumn={() => handleAddColumn(col)}
                      onDeleteColumn={() => handleDeleteColumn(col)}
                      onTabNavigation={handleTabNavigation}
                    />
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      </div>

      {showFunctionTester && (
        <FunctionTester
          onClose={() => setShowFunctionTester(false)}
          mathFunctions={mathFunctions}
          dataFunctions={dataFunctions}
        />
      )}
    </div>
  );
}

export default SheetApp;
