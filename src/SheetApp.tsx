import React, { useState, useCallback, useRef, useEffect } from 'react';
import Toolbar from './components/Toolbar';
import FormulaBar from './components/FormulaBar';
import Cell from './components/Cell';
import FunctionTester from './components/FunctionTester';
import './SheetApp.css';
import * as XLSX from 'xlsx';
import {
  CellStyle,
  CellData,
  SelectionRange,
  CellPosition,
  ResizeLinePosition,
  MathFunctions,
  DataFunctions,
  CellDimensions
} from './types';

const DEFAULT_ROWS = 50;
const DEFAULT_COLS = 26;

// Helper functions for cell references and formulas
const getCellName = (row: number, col: number): string => 
  String.fromCharCode(65 + col) + String(row + 1);

const parseCell = (cellStr: string): CellPosition => {
  const col = cellStr.charCodeAt(0) - 65;
  const row = parseInt(cellStr.slice(1), 10) - 1;
  return { row, col };
};

const parseRange = (rangeStr: string): [string, string] => {
  const parts = rangeStr.split(':');
  return parts.length === 1 ? [parts[0], parts[0]] : [parts[0], parts[1]];
};

const SheetApp: React.FC = () => {
  // State management
  const [numRows, setNumRows] = useState<number>(DEFAULT_ROWS);
  const [numCols, setNumCols] = useState<number>(DEFAULT_COLS);
  const [cells, setCells] = useState<Record<string, CellData>>({});
  const [cellStyles, setCellStyles] = useState<Record<string, CellStyle>>({});
  const [selectedCell, setSelectedCell] = useState<string>('A1');
  const [selectionRange, setSelectionRange] = useState<SelectionRange | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [dragStart, setDragStart] = useState<string | null>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const [fontFamily, setFontFamily] = useState<string>('Arial');
  const [dragSource, setDragSource] = useState<string | null>(null);
  const [fillSource, setFillSource] = useState<string | null>(null);
  const [fillTarget, setFillTarget] = useState<string | null>(null);
  const [formulaText, setFormulaText] = useState<string | null>(null);
  const [columnWidths, setColumnWidths] = useState<CellDimensions>({});
  const [rowHeights, setRowHeights] = useState<CellDimensions>({});
  const [isResizing, setIsResizing] = useState<boolean>(false);
  const resizeStartRef = useRef<number | null>(null);
  const resizeTypeRef = useRef<'row' | 'column' | null>(null);
  const resizeIndexRef = useRef<number | null>(null);
  const [resizeLinePosition, setResizeLinePosition] = useState<ResizeLinePosition | null>(null);
  const [showFunctionTester, setShowFunctionTester] = useState<boolean>(false);

  // Cell style management
  const updateCellStyle = (cellName: string, styleUpdate: Partial<CellStyle>): void => {
    setCellStyles(prev => ({
      ...prev,
      [cellName]: { 
        ...(prev[cellName] || {}), 
        ...styleUpdate 
      }
    }));
  };

  // Mathematical functions
  const mathFunctions: MathFunctions = {
    SUM: (values) => values.reduce((a, b) => a + b, 0),
    AVERAGE: (values) => values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0,
    MAX: (values) => values.length ? Math.max(...values) : 0,
    MIN: (values) => values.length ? Math.min(...values) : 0,
    COUNT: (values) => values.filter(value => {
      const num = parseFloat(String(value));
      return !isNaN(num) && isFinite(num);
    }).length
  };

  // Data quality functions
  const dataFunctions: DataFunctions = {
    TRIM: ({ value }) => {
      if (Array.isArray(value)) return value[0].trim();
      return value.trim();
    },
    UPPER: ({ value }) => {
      if (Array.isArray(value)) return value[0].toUpperCase();
      return value.toUpperCase();
    },
    LOWER: ({ value }) => {
      if (Array.isArray(value)) return value[0].toLowerCase();
      return value.toLowerCase();
    },
    REMOVE_DUPLICATES: ({ value }) => {
      if (!Array.isArray(value)) return value;
      
      const stringifyRow = (row: string | string[]): string => {
        if (Array.isArray(row)) {
          return row.join('|');
        }
        return String(row);
      };

      const uniqueRows = new Set(value.map(stringifyRow));
      
      return Array.from(uniqueRows).map(str => {
        if (str.includes('|')) {
          return str.split('|');
        }
        return str;
      }) as Array<string | string[]>;
    },
    FIND_AND_REPLACE: ({ value, find, replace }) => {
      if (!value || !find) return value;
      
      if (Array.isArray(value)) {
        return value.map(val => 
          String(val).replace(new RegExp(find, 'g'), replace || '')
        );
      }
      
      return String(value).replace(new RegExp(find, 'g'), replace || '');
    }
  };

  // Selection handling
  const handleCellSelect = (cellName: string, e?: React.MouseEvent): void => {
    if (e?.shiftKey && selectedCell) {
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
  const handleMouseDown = (cellName: string, e: React.MouseEvent): void => {
    if (!e.shiftKey) {
      setIsDragging(true);
      setDragStart(cellName);
      setSelectedCell(cellName);
      setSelectionRange(null);
    }
  };

  const handleMouseMove = (cellName: string): void => {
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

  const handleMouseUp = (): void => {
    setIsDragging(false);
    setDragStart(null);
  };

  // Effect for mouse up handler
  useEffect(() => {
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  const handleBold = (): void => {
    if (selectedCell) {
      updateCellStyle(selectedCell, {
        bold: !cellStyles[selectedCell]?.bold
      });
    }
  };

  const handleItalic = (): void => {
    if (selectedCell) {
      updateCellStyle(selectedCell, {
        italic: !cellStyles[selectedCell]?.italic
      });
    }
  };

  const handleFontSize = (size: string): void => {
    if (selectedCell) {
      updateCellStyle(selectedCell, { fontSize: size });
    }
  };

  const handleColor = (color: string): void => {
    if (selectedCell) {
      updateCellStyle(selectedCell, { color });
    }
  };

  const handleFontFamily = (font: string): void => {
    if (selectedCell) {
      updateCellStyle(selectedCell, { fontFamily: font });
      setFontFamily(font);
    }
  };

  const handleStrikethrough = (): void => {
    if (selectedCell) {
      updateCellStyle(selectedCell, {
        strikethrough: !cellStyles[selectedCell]?.strikethrough
      });
    }
  };

  const handleAlignment = (alignment: 'left' | 'center' | 'right'): void => {
    if (selectedCell) {
      updateCellStyle(selectedCell, { textAlign: alignment });
    }
  };

  const handleSave = (): void => {
    // Implementation needed
  };

  const handleLoad = (): void => {
    // Implementation needed
  };

  const handleCellChange = (cellName: string, value: string): void => {
    // Implementation needed
  };

  const handleFormulaSelect = (formula: string): void => {
    if (selectedCell) {
      handleCellChange(selectedCell, `=${formula}`);
    }
  };

  return (
    <div className="sheet-app">
      <Toolbar
        selectedCell={selectedCell}
        cellStyles={cellStyles}
        currentFont={fontFamily}
        onBold={handleBold}
        onItalic={handleItalic}
        onFontSize={handleFontSize}
        onColor={handleColor}
        onFontFamily={handleFontFamily}
        onFormulaSelect={handleFormulaSelect}
        onAlignment={handleAlignment}
        onOpenFunctionTester={() => setShowFunctionTester(true)}
        onSave={handleSave}
        onLoad={handleLoad}
      />
      <FormulaBar
        selectedCell={selectedCell}
        value={cells[selectedCell]?.formula || cells[selectedCell]?.value || ''}
        onChange={handleCellChange}
        formulaText={formulaText}
        onCellSelect={handleCellSelect}
      />
      <div className="grid-container" ref={gridRef}>
        <table className="spreadsheet-grid">
          <thead>
            <tr>
              <th className="corner-header"></th>
              {Array.from({ length: numCols }, (_, colIndex) => (
                <th key={colIndex} className="column-header">
                  {String.fromCharCode(65 + colIndex)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: numRows }, (_, rowIndex) => (
              <tr key={rowIndex}>
                <td className="row-header">{rowIndex + 1}</td>
                {Array.from({ length: numCols }, (_, colIndex) => {
                  const cellName = getCellName(rowIndex, colIndex);
                  const cellData = cells[cellName] || { value: '', formula: '' };
                  const isInSelection = selectionRange && (
                    rowIndex >= selectionRange.startRow &&
                    rowIndex <= selectionRange.endRow &&
                    colIndex >= selectionRange.startCol &&
                    colIndex <= selectionRange.endCol
                  );
                  
                  let selectionClass = null;
                  if (cellName === selectedCell) {
                    selectionClass = 'selected-cell primary-selected';
                  } else if (isInSelection) {
                    selectionClass = 'selected-cell';
                  }

                  return (
                    <Cell
                      key={cellName}
                      name={cellName}
                      value={cellData.formula || cellData.value}
                      displayValue={cellData.value}
                      onChange={(value) => handleCellChange(cellName, value)}
                      onSelect={(e) => handleCellSelect(cellName, e)}
                      onMouseDown={(e) => handleMouseDown(cellName, e)}
                      onMouseMove={() => handleMouseMove(cellName)}
                      isSelected={selectionClass}
                      styles={cellStyles[cellName]}
                      onDragStart={(name) => setDragSource(name)}
                      onDragEnd={() => setDragSource(null)}
                      onFillDragStart={(name) => setFillSource(name)}
                      onFillDragMove={(e) => {
                        if (fillSource) {
                          const rect = (e.target as HTMLElement).getBoundingClientRect();
                          setFillTarget(cellName);
                        }
                      }}
                      onFillDragEnd={() => {
                        setFillSource(null);
                        setFillTarget(null);
                      }}
                      onAddRow={() => setNumRows(prev => prev + 1)}
                      onDeleteRow={() => setNumRows(prev => Math.max(1, prev - 1))}
                      onAddColumn={() => setNumCols(prev => prev + 1)}
                      onDeleteColumn={() => setNumCols(prev => Math.max(1, prev - 1))}
                      onTabNavigation={(name, shiftKey, isEnter) => {
                        const current = parseCell(name);
                        let nextCell;
                        
                        if (isEnter) {
                          // Move down on enter
                          if (current.row < numRows - 1) {
                            nextCell = getCellName(current.row + 1, current.col);
                          }
                        } else if (shiftKey) {
                          // Move left on shift+tab
                          if (current.col > 0) {
                            nextCell = getCellName(current.row, current.col - 1);
                          } else if (current.row > 0) {
                            nextCell = getCellName(current.row - 1, numCols - 1);
                          }
                        } else {
                          // Move right on tab
                          if (current.col < numCols - 1) {
                            nextCell = getCellName(current.row, current.col + 1);
                          } else if (current.row < numRows - 1) {
                            nextCell = getCellName(current.row + 1, 0);
                          }
                        }
                        
                        if (nextCell) {
                          handleCellSelect(nextCell);
                        }
                      }}
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
          onSelect={handleFormulaSelect}
        />
      )}
    </div>
  );
};

export default SheetApp; 