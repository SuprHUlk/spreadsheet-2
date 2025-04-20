import React, { useState, useCallback, useRef, useEffect } from "react";
import Toolbar from "./components/Toolbar"; // Fixed import without .tsx extension
import FormulaBar from "./components/FormulaBar";
import Cell from "./components/Cell";
import FunctionTester from "./components/FunctionTester";
import "./SheetApp.css";
import * as XLSX from "xlsx";
import { useAppDispatch, useAppSelector } from "./redux/hooks";
import {
  updateCell,
  updateCellStyle,
  undo,
  redo,
  setLastSaved,
} from "./redux/slices/spreadsheetSlice";

const DEFAULT_ROWS = 50;
const DEFAULT_COLS = 26;

// Types
interface CellPosition {
  row: number;
  col: number;
}

interface SelectionRange {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

interface CellStyle {
  fontWeight?: string;
  fontStyle?: string;
  fontSize?: string;
  color?: string;
  fontFamily?: string;
  textDecoration?: string;
  textAlign?: string;
}

interface ResizeStartData {
  position: number;
  size: number;
  initialPosition: number;
}

type ColumnWidthsState = Record<number, number>;
type RowHeightsState = Record<number, number>;

type MathFunction = (values: number[]) => number;
interface MathFunctions {
  [key: string]: MathFunction;
}

type DataFunction = (...args: any[]) => any;
interface DataFunctions {
  [key: string]: DataFunction;
}

// Helper functions for cell references and formulas
const getCellName = (row: number, col: number): string =>
  String.fromCharCode(65 + col) + (row + 1);

const parseCell = (cellStr: string): CellPosition => {
  const col = cellStr.charCodeAt(0) - 65;
  const row = parseInt(cellStr.slice(1), 10) - 1;
  return { row, col };
};

const parseRange = (rangeStr: string): [string, string] => {
  const parts = rangeStr.split(":");
  return parts.length === 1 ? [parts[0], parts[0]] : [parts[0], parts[1]];
};

const SheetApp: React.FC = () => {
  // Redux dispatch
  const dispatch = useAppDispatch();

  // Get spreadsheet state from Redux
  const { cells, cellStyles } = useAppSelector((state) => state.spreadsheet);

  // Local state for UI elements
  const [numRows, setNumRows] = useState<number>(DEFAULT_ROWS);
  const [numCols, setNumCols] = useState<number>(DEFAULT_COLS);
  const [selectedCell, setSelectedCell] = useState<string>("A1");
  const [selectionRange, setSelectionRange] = useState<SelectionRange | null>(
    null
  );
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [dragStart, setDragStart] = useState<string | null>(null);
  const gridRef = useRef<HTMLDivElement | null>(null);
  const [fontFamily, setFontFamily] = useState<string>("Arial");
  const [dragSource, setDragSource] = useState<string | null>(null);
  const [fillSource, setFillSource] = useState<string | null>(null);
  const [fillTarget, setFillTarget] = useState<string | null>(null);
  const [formulaText, setFormulaText] = useState<string | null>(null);
  const [columnWidths, setColumnWidths] = useState<ColumnWidthsState>({});
  const [rowHeights, setRowHeights] = useState<RowHeightsState>({});
  const [isResizing, setIsResizing] = useState<boolean>(false);
  const resizeStartRef = useRef<ResizeStartData | null>(null);
  const resizeTypeRef = useRef<"column" | "row" | null>(null);
  const resizeIndexRef = useRef<number | null>(null);
  const [resizeLinePosition, setResizeLinePosition] = useState<number | null>(
    null
  );
  const [showFunctionTester, setShowFunctionTester] = useState<boolean>(false);

  // Cell style management - now uses Redux
  const updateCellStyleRedux = (
    cellName: string,
    styleProperty: string,
    value: any
  ): void => {
    dispatch(updateCellStyle({ cellName, styleProperty, value }));
  };

  // Mathematical functions
  const mathFunctions: MathFunctions = {
    SUM: (values) => values.reduce((a, b) => a + b, 0),
    AVERAGE: (values) =>
      values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0,
    MAX: (values) => (values.length ? Math.max(...values) : 0),
    MIN: (values) => (values.length ? Math.min(...values) : 0),
    COUNT: (values) =>
      values.filter((value) => {
        const num = parseFloat(String(value));
        return !isNaN(num) && isFinite(num);
      }).length,
  };

  // Data quality functions
  const dataFunctions: DataFunctions = {
    TRIM: (value: string | string[]) => {
      if (Array.isArray(value)) return value[0].trim();
      return value.trim();
    },
    UPPER: (value: string | string[]) => {
      if (Array.isArray(value)) return value[0].toUpperCase();
      return value.toUpperCase();
    },
    LOWER: (value: string | string[]) => {
      if (Array.isArray(value)) return value[0].toLowerCase();
      return value.toLowerCase();
    },
    // REMOVE_DUPLICATES: (range: any[] | any) => {
    //   if (!Array.isArray(range)) return range;

    //   // Convert range to string for comparison
    //   const stringifyRow = (row: any): string => {
    //     if (Array.isArray(row)) {
    //       return row.join("|");
    //     }
    //     return String(row);
    //   };

    //   // Create a Set of stringified rows to remove duplicates
    //   const uniqueRows = new Set(range.map(stringifyRow));

    //   // Convert back to original format
    //   return Array.from(uniqueRows).map((str) => {
    //     if (typeof str === "string" && str.includes("|")) {
    //       return str.split("|");
    //     }
    //     return str;
    //   });
    // },
    // FIND_AND_REPLACE: (
    //   text: string | string[],
    //   find: string,
    //   replace?: string
    // ) => {
    //   if (!text || !find) return text;

    //   // Handle array of values (range of cells)
    //   if (Array.isArray(text)) {
    //     return text.map((value) =>
    //       String(value).replace(new RegExp(find, "g"), replace || "")
    //     );
    //   }

    //   return String(text).replace(new RegExp(find, "g"), replace || "");
    // },
  };

  // Selection handling
  const handleCellSelect = (cellName: string, e?: React.MouseEvent): void => {
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
    window.addEventListener("mouseup", handleMouseUp);
    return () => window.removeEventListener("mouseup", handleMouseUp);
  }, []);

  // Function to update cell references in a formula
  const updateFormulaReferences = useCallback(
    (formula: string | undefined, fromCell: string, toCell: string): string => {
      if (!formula?.startsWith("=")) return formula || "";

      const expression = formula.slice(1);
      return (
        "=" +
        expression.replace(/[A-Z]+[0-9]+/g, (ref) => {
          if (ref === fromCell) return toCell;
          return ref;
        })
      );
    },
    []
  );

  // Function to adjust formula for fill operation
  const adjustFormulaForFill = useCallback(
    (
      formula: string | undefined,
      sourceCell: string,
      targetCell: string
    ): string => {
      if (!formula?.startsWith("=")) return formula || "";

      const source = parseCell(sourceCell);
      const target = parseCell(targetCell);
      const rowDiff = target.row - source.row;
      const colDiff = target.col - source.col;

      const expression = formula.slice(1);
      return (
        "=" +
        expression.replace(/[A-Z]+[0-9]+/g, (ref) => {
          const refCell = parseCell(ref);
          const newRow = refCell.row + rowDiff;
          const newCol = refCell.col + colDiff;

          // Check if the new reference is valid
          if (
            newRow >= 0 &&
            newRow < numRows &&
            newCol >= 0 &&
            newCol < numCols
          ) {
            return getCellName(newRow, newCol);
          }
          return ref; // Keep original reference if new one would be invalid
        })
      );
    },
    [numRows, numCols]
  );

  // Formula evaluation
  const evaluateFormula = useCallback(
    (formula: string, cellName: string): string | number => {
      if (!formula.startsWith("=")) return formula;

      const expression = formula.slice(1);
      const funcMatch = expression.match(/(\w+)\((.*)\)/);

      if (!funcMatch) {
        // Handle basic arithmetic
        try {
          const evalExpr = expression.replace(/[A-Z]\d+/g, (ref) => {
            if (ref === cellName) return "0"; // Prevent circular references
            const value = getCellDisplayValue(ref);
            return isNaN(parseFloat(value))
              ? "0"
              : parseFloat(value).toString();
          });
          // Using Function constructor to evaluate mathematical expressions
          return Function('"use strict";return (' + evalExpr + ")")();
        } catch (e) {
          return "#ERROR!";
        }
      }

      const [_, func, args] = funcMatch;
      const functionName = func.toUpperCase();

      if (mathFunctions[functionName]) {
        const [startRef, endRef] = parseRange(args);
        const start = parseCell(startRef);
        const end = parseCell(endRef);
        const values: number[] = [];

        for (
          let row = Math.min(start.row, end.row);
          row <= Math.max(start.row, end.row);
          row++
        ) {
          for (
            let col = Math.min(start.col, end.col);
            col <= Math.max(start.col, end.col);
            col++
          ) {
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
        const argValues = args.split(",").map((arg) => {
          const trimmed = arg.trim();
          if (trimmed.match(/^[A-Z]\d+$/)) {
            return getCellDisplayValue(trimmed);
          }
          return trimmed.replace(/^['"]|['"]$/g, "");
        });

        return dataFunctions[functionName](...argValues);
      }

      return "#UNKNOWN_FUNCTION";
    },
    [cells]
  );

  const getCellDisplayValue = useCallback(
    (cellName: string): string => {
      const value = cells[cellName] || "";
      if (value.startsWith("=")) {
        try {
          const result = evaluateFormula(value, cellName);
          // Format numbers to avoid unnecessary decimal places
          if (typeof result === "number") {
            return Number.isInteger(result)
              ? result.toString()
              : result.toFixed(2);
          }
          return String(result);
        } catch (e) {
          return "#ERROR!";
        }
      }
      return value;
    },
    [cells, evaluateFormula]
  );

  // Cell operations - updated to use Redux
  const handleCellChange = (cellName: string, value: string): void => {
    dispatch(updateCell({ cellName, value }));
  };

  // Row and column operations
  const handleAddRow = useCallback(
    (index: number): void => {
      // Update cells state to accommodate new row
      const newCells = { ...cells };
      // Shift all cells below the insertion point up by one row
      Object.entries(newCells).forEach(([cellName, value]) => {
        const { row, col } = parseCell(cellName);
        if (row > index) {
          const newCellName = getCellName(row + 1, col);
          dispatch(updateCell({ cellName: newCellName, value }));
          // Delete the old cell through Redux in the next update
          dispatch(updateCell({ cellName, value: "" }));
        }
      });
      setNumRows(numRows + 1);
    },
    [cells, numRows, dispatch]
  );

  const handleDeleteRow = useCallback(
    (index: number): void => {
      if (numRows > 1) {
        const newCells = { ...cells };
        // Remove cells in the deleted row and shift cells up
        Object.entries(newCells).forEach(([cellName, value]) => {
          const { row, col } = parseCell(cellName);
          if (row === index) {
            // Delete cell by setting its value to empty string
            dispatch(updateCell({ cellName, value: "" }));
          } else if (row > index) {
            // Create new cell in the shifted position
            const newCellName = getCellName(row - 1, col);
            dispatch(updateCell({ cellName: newCellName, value }));
            // Delete the old cell
            dispatch(updateCell({ cellName, value: "" }));
          }
        });
        setNumRows(numRows - 1);
      }
    },
    [cells, numRows, dispatch]
  );

  const handleAddColumn = useCallback(
    (index: number): void => {
      const newCells = { ...cells };
      // Shift all cells to the right of the insertion point by one column
      Object.entries(newCells).forEach(([cellName, value]) => {
        const { row, col } = parseCell(cellName);
        if (col > index) {
          const newCellName = getCellName(row, col + 1);
          dispatch(updateCell({ cellName: newCellName, value }));
          // Delete the old cell
          dispatch(updateCell({ cellName, value: "" }));
        }
      });
      setNumCols(numCols + 1);
    },
    [cells, numCols, dispatch]
  );

  const handleDeleteColumn = useCallback(
    (index: number): void => {
      if (numCols > 1) {
        const newCells = { ...cells };
        // Remove cells in the deleted column and shift cells left
        Object.entries(newCells).forEach(([cellName, value]) => {
          const { row, col } = parseCell(cellName);
          if (col === index) {
            // Delete cell by setting empty value
            dispatch(updateCell({ cellName, value: "" }));
          } else if (col > index) {
            // Create new cell in the shifted position
            const newCellName = getCellName(row, col - 1);
            dispatch(updateCell({ cellName: newCellName, value }));
            // Delete the old cell
            dispatch(updateCell({ cellName, value: "" }));
          }
        });
        setNumCols(numCols - 1);
      }
    },
    [cells, numCols, dispatch]
  );

  // Get cell class based on selection state
  const getCellClass = (row: number, col: number): string => {
    const cellName = getCellName(row, col);

    if (cellName === selectedCell) {
      return "selected primary";
    }

    if (selectionRange) {
      const { startRow, endRow, startCol, endCol } = selectionRange;
      if (
        row >= startRow &&
        row <= endRow &&
        col >= startCol &&
        col <= endCol
      ) {
        return "selected";
      }
    }

    return "";
  };

  // Style handlers - updated to use Redux
  const handleBold = (): void => {
    if (!selectedCell) return;
    const currentBold = cellStyles[selectedCell]?.bold || false;
    updateCellStyleRedux(selectedCell, "bold", !currentBold);
  };

  const handleItalic = (): void => {
    if (!selectedCell) return;
    const currentItalic = cellStyles[selectedCell]?.italic || false;
    updateCellStyleRedux(selectedCell, "italic", !currentItalic);
  };

  const handleFontSize = (size: string): void => {
    if (!selectedCell) return;
    updateCellStyleRedux(selectedCell, "fontSize", size);
  };

  const handleColor = (color: string): void => {
    if (!selectedCell) return;
    updateCellStyleRedux(selectedCell, "textColor", color);
  };

  const handleFontFamily = (font: string): void => {
    if (!selectedCell) return;
    setFontFamily(font);
    updateCellStyleRedux(selectedCell, "fontFamily", font);
  };

  const handleStrikethrough = (): void => {
    if (!selectedCell) return;
    const currentDecoration =
      cellStyles[selectedCell]?.textDecoration === "line-through";
    updateCellStyleRedux(
      selectedCell,
      "textDecoration",
      currentDecoration ? "none" : "line-through"
    );
  };

  const handleAlignment = (alignment: string): void => {
    if (!selectedCell) return;
    updateCellStyleRedux(
      selectedCell,
      "textAlign",
      alignment as "left" | "center" | "right"
    );
  };

  // New undo/redo handlers
  const handleUndo = (): void => {
    dispatch(undo());
  };

  const handleRedo = (): void => {
    dispatch(redo());
  };

  // Drag and drop handlers
  const handleCellDragStart = (cellName: string): void => {
    setDragSource(cellName);
    setSelectedCell(cellName);
  };

  const handleCellDragEnd = (): void => {
    if (dragSource && selectedCell && dragSource !== selectedCell) {
      // Move cell content
      const sourceContent = cells[dragSource];
      const sourceStyle = cellStyles[dragSource];
      const targetContent = cells[selectedCell];
      const targetStyle = cellStyles[selectedCell];

      // Update cells using Redux
      const newCells = { ...cells };

      // Update the moved cell content and adjust any formulas
      newCells[selectedCell] = updateFormulaReferences(
        sourceContent,
        dragSource,
        selectedCell
      );
      newCells[dragSource] = updateFormulaReferences(
        targetContent,
        selectedCell,
        dragSource
      );

      // Update any formulas that reference the moved cells
      Object.entries(newCells).forEach(([cellName, value]) => {
        if (value?.startsWith("=")) {
          newCells[cellName] = updateFormulaReferences(
            updateFormulaReferences(value, dragSource, selectedCell),
            selectedCell,
            dragSource
          );
        }
      });

      // Dispatch multiple cell updates to Redux
      Object.entries(newCells).forEach(([cellName, value]) => {
        if (cells[cellName] !== value) {
          dispatch(updateCell({ cellName, value }));
        }
      });

      // Update cell styles
      Object.entries(sourceStyle || {}).forEach(([property, value]) => {
        updateCellStyleRedux(selectedCell, property, value);
      });

      Object.entries(targetStyle || {}).forEach(([property, value]) => {
        updateCellStyleRedux(dragSource, property, value);
      });
    }
    setDragSource(null);
  };

  // Fill handle handlers
  const handleFillDragStart = (cellName: string): void => {
    setFillSource(cellName);
    setSelectedCell(cellName);
  };

  const getCellFromPoint = (x: number, y: number): string | null => {
    const table = gridRef.current;
    if (!table) return null;

    const cells = table.getElementsByTagName("td");
    const tableRect = table.getBoundingClientRect();

    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      const cellRect = cell.getBoundingClientRect();
      if (
        x >= cellRect.left &&
        x <= cellRect.right &&
        y >= cellRect.top &&
        y <= cellRect.bottom &&
        x >= tableRect.left &&
        x <= tableRect.right &&
        y >= tableRect.top &&
        y <= tableRect.bottom
      ) {
        return cell.getAttribute("data-cell-name");
      }
    }
    return null;
  };

  const handleFillDragMove = (e: React.MouseEvent): void => {
    const targetCell = getCellFromPoint(e.clientX, e.clientY);
    if (targetCell && fillSource) {
      setFillTarget(targetCell);

      const start = parseCell(fillSource);
      const end = parseCell(targetCell);

      setSelectionRange({
        startRow: Math.min(start.row, end.row),
        endRow: Math.max(start.row, end.row),
        startCol: Math.min(start.col, end.col),
        endCol: Math.max(start.col, end.col),
      });
    }
  };

  const handleFillDragEnd = (): void => {
    if (fillSource && fillTarget) {
      const sourceValue = cells[fillSource];
      const sourceStyle = cellStyles[fillSource];

      if (typeof sourceValue === "string" && sourceValue.match(/^\d+$/)) {
        // Numeric sequence
        const startNum = parseInt(sourceValue);
        const start = parseCell(fillSource);
        const end = parseCell(fillTarget);

        const isVertical = start.col === end.col;
        let counter = 0;

        if (isVertical) {
          for (
            let row = Math.min(start.row, end.row);
            row <= Math.max(start.row, end.row);
            row++
          ) {
            const cellName = getCellName(row, start.col);
            if (cellName !== fillSource) {
              dispatch(
                updateCell({ cellName, value: String(startNum + counter) })
              );

              // Apply styles from source cell
              if (sourceStyle) {
                Object.entries(sourceStyle).forEach(([property, value]) => {
                  updateCellStyleRedux(cellName, property, value);
                });
              }
              counter++;
            }
          }
        } else {
          for (
            let col = Math.min(start.col, end.col);
            col <= Math.max(start.col, end.col);
            col++
          ) {
            const cellName = getCellName(start.row, col);
            if (cellName !== fillSource) {
              dispatch(
                updateCell({ cellName, value: String(startNum + counter) })
              );

              // Apply styles from source cell
              if (sourceStyle) {
                Object.entries(sourceStyle).forEach(([property, value]) => {
                  updateCellStyleRedux(cellName, property, value);
                });
              }
              counter++;
            }
          }
        }
      } else if (sourceValue?.startsWith("=")) {
        // Formula fill - adjust references
        const start = parseCell(fillSource);
        const end = parseCell(fillTarget);

        for (
          let row = Math.min(start.row, end.row);
          row <= Math.max(start.row, end.row);
          row++
        ) {
          for (
            let col = Math.min(start.col, end.col);
            col <= Math.max(start.col, end.col);
            col++
          ) {
            const cellName = getCellName(row, col);
            if (cellName !== fillSource) {
              const adjustedFormula = adjustFormulaForFill(
                sourceValue,
                fillSource,
                cellName
              );
              dispatch(updateCell({ cellName, value: adjustedFormula }));

              // Apply styles from source cell
              if (sourceStyle) {
                Object.entries(sourceStyle).forEach(([property, value]) => {
                  updateCellStyleRedux(cellName, property, value);
                });
              }
            }
          }
        }
      } else {
        // Copy value
        const start = parseCell(fillSource);
        const end = parseCell(fillTarget);

        for (
          let row = Math.min(start.row, end.row);
          row <= Math.max(start.row, end.row);
          row++
        ) {
          for (
            let col = Math.min(start.col, end.col);
            col <= Math.max(start.col, end.col);
            col++
          ) {
            const cellName = getCellName(row, col);
            if (cellName !== fillSource) {
              dispatch(updateCell({ cellName, value: sourceValue }));

              // Apply styles from source cell
              if (sourceStyle) {
                Object.entries(sourceStyle).forEach(([property, value]) => {
                  updateCellStyleRedux(cellName, property, value);
                });
              }
            }
          }
        }
      }
    }

    setFillSource(null);
    setFillTarget(null);
    setSelectionRange(null);
  };

  // Add this handler function inside the SheetApp component
  const handleFormulaSelect = (formula: string): void => {
    if (!selectedCell || !formula) return;
    setFormulaText(formula);
  };

  // Add these new handlers
  const handleResizeStart = (
    e: React.MouseEvent,
    type: "column" | "row",
    index: number
  ): void => {
    e.preventDefault();
    e.stopPropagation();
    setIsResizing(true);

    const rect = (
      e.currentTarget.closest("th") as HTMLElement
    ).getBoundingClientRect();
    const tableRect = gridRef.current!.getBoundingClientRect();

    let initialPosition;
    if (type === "column") {
      initialPosition = e.clientX;
    } else {
      initialPosition = e.clientY;
    }

    resizeStartRef.current = {
      position: type === "column" ? e.clientX : e.clientY,
      size:
        type === "column"
          ? columnWidths[index] || 100
          : rowHeights[index] || 24,
      initialPosition,
    };
    resizeTypeRef.current = type;
    resizeIndexRef.current = index;
    setResizeLinePosition(initialPosition);

    document.body.classList.add("resizing");
    if (type === "row") {
      document.body.classList.add("row");
    }
  };

  const handleResizeMove = useCallback(
    (e: MouseEvent): void => {
      if (!isResizing || !resizeStartRef.current) return;

      const startData = resizeStartRef.current;
      const currentPos =
        resizeTypeRef.current === "column" ? e.clientX : e.clientY;

      setResizeLinePosition(currentPos);
    },
    [isResizing]
  );

  useEffect(() => {
    if (isResizing) {
      const handleMove = (e: MouseEvent): void => {
        e.preventDefault();
        e.stopPropagation();
        handleResizeMove(e);
      };

      const handleUp = (e: MouseEvent): void => {
        e.preventDefault();
        e.stopPropagation();

        const type = resizeTypeRef.current;
        const index = resizeIndexRef.current;
        const startData = resizeStartRef.current;

        if (!type || index === null || !startData) return;

        const endPos = type === "column" ? e.clientX : e.clientY;
        const diff = endPos - startData.position;

        if (type === "column") {
          const newWidth = Math.max(50, startData.size + diff);
          setColumnWidths((prev) => ({
            ...prev,
            [index]: newWidth,
          }));
        } else {
          const newHeight = Math.max(24, startData.size + diff);
          setRowHeights((prev) => ({
            ...prev,
            [index]: newHeight,
          }));
        }

        setIsResizing(false);
        setResizeLinePosition(null);
        resizeStartRef.current = null;
        resizeTypeRef.current = null;
        resizeIndexRef.current = null;
        document.body.classList.remove("resizing", "row");
      };

      window.addEventListener("mousemove", handleMove, { capture: true });
      window.addEventListener("mouseup", handleUp, { capture: true });

      return () => {
        window.removeEventListener("mousemove", handleMove, { capture: true });
        window.removeEventListener("mouseup", handleUp, { capture: true });
      };
    }
  }, [isResizing, handleResizeMove]);

  // Save spreadsheet data - updated to record last saved time in Redux
  const handleSave = (): void => {
    // Create a new workbook
    const wb = XLSX.utils.book_new();

    // Convert cells data to array format for Excel
    const data: any[][] = [];
    let maxRow = 0;
    let maxCol = 0;

    // Find the maximum used row and column
    Object.keys(cells).forEach((cellName) => {
      const { row, col } = parseCell(cellName);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    });

    // Create data array with all cells (including empty ones)
    for (let row = 0; row <= maxRow; row++) {
      const rowData: any[] = [];
      for (let col = 0; col <= maxCol; col++) {
        const cellName = getCellName(row, col);
        rowData.push(getCellDisplayValue(cellName));
      }
      data.push(rowData);
    }

    // Create worksheet from data
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Add column widths
    ws["!cols"] = Array(maxCol + 1)
      .fill(null)
      .map((_, i) => ({
        wch: columnWidths[i] ? Math.floor(columnWidths[i] / 7) : 14, // Convert pixels to Excel width units (approx)
      }));

    // Add row heights
    ws["!rows"] = Array(maxRow + 1)
      .fill(null)
      .map((_, i) => ({
        hpt: rowHeights[i] ? Math.floor(rowHeights[i] * 0.75) : 18, // Convert pixels to Excel height units (approx)
      }));

    // Add styles
    Object.keys(cellStyles).forEach((cellName) => {
      const style = cellStyles[cellName];
      const addr = XLSX.utils.decode_cell(cellName);
      const cell = ws[XLSX.utils.encode_cell(addr)];
      if (cell) {
        cell.s = {
          font: {
            bold: style.bold,
            italic: style.italic,
            name: style.fontFamily,
            sz: parseInt(style.fontSize || "11"),
            color: { rgb: style.textColor?.replace("#", "") },
          },
        };
      }
    });

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    // Generate Excel file and trigger download
    XLSX.writeFile(wb, "spreadsheet.xlsx");

    // Update last saved timestamp in Redux
    dispatch(setLastSaved());
  };

  // Load spreadsheet data
  const handleLoad = (): void => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx,.xls";
    input.onchange = (e: Event) => {
      const target = e.target as HTMLInputElement;
      const file = target.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (event: ProgressEvent<FileReader>) => {
          try {
            const data = new Uint8Array(event.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });

            // Get the first worksheet
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            // Convert worksheet to array of arrays
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
            }) as any[][];

            // Clear existing data by clearing Redux state for each existing cell
            Object.keys(cells).forEach((cellName) => {
              dispatch(updateCell({ cellName, value: "" }));
            });

            // Update dimensions
            setNumRows(Math.max(DEFAULT_ROWS, jsonData.length));
            setNumCols(
              Math.max(
                DEFAULT_COLS,
                Math.max(...jsonData.map((row) => row.length || 0))
              )
            );

            // Convert data to our cell format and update Redux
            jsonData.forEach((row, rowIndex) => {
              row.forEach((cellValue, colIndex) => {
                if (cellValue !== null && cellValue !== undefined) {
                  const cellName = getCellName(rowIndex, colIndex);
                  dispatch(
                    updateCell({ cellName, value: cellValue.toString() })
                  );

                  // Extract cell style if available
                  const cell =
                    worksheet[
                      XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
                    ];
                  if (cell?.s) {
                    if (cell.s.font) {
                      const font = cell.s.font as any;
                      if (font.bold)
                        updateCellStyleRedux(cellName, "fontWeight", "bold");
                      if (font.italic)
                        updateCellStyleRedux(cellName, "fontStyle", "italic");
                      if (font.name)
                        updateCellStyleRedux(cellName, "fontFamily", font.name);
                      if (font.sz)
                        updateCellStyleRedux(
                          cellName,
                          "fontSize",
                          `${font.sz}px`
                        );
                      if (font.color?.rgb)
                        updateCellStyleRedux(
                          cellName,
                          "color",
                          `#${font.color.rgb}`
                        );
                    }
                  }
                }
              });
            });

            // Update column widths
            const newColumnWidths: ColumnWidthsState = {};
            if (worksheet["!cols"]) {
              worksheet["!cols"].forEach((col: any, index: number) => {
                if (col?.wch) {
                  newColumnWidths[index] = col.wch * 7; // Convert Excel width units to pixels (approx)
                }
              });
            }

            // Update row heights
            const newRowHeights: RowHeightsState = {};
            if (worksheet["!rows"]) {
              worksheet["!rows"].forEach((row: any, index: number) => {
                if (row?.hpt) {
                  newRowHeights[index] = row.hpt / 0.75; // Convert Excel height units to pixels (approx)
                }
              });
            }

            setColumnWidths(newColumnWidths);
            setRowHeights(newRowHeights);
          } catch (error) {
            console.error("Error loading spreadsheet:", error);
            alert("Error loading spreadsheet: Invalid file format");
          }
        };
        reader.readAsArrayBuffer(file);
      }
    };
    input.click();
  };

  // Add this function to handle tab navigation
  const handleTabNavigation = (
    cellName: string,
    isShiftKey: boolean,
    isEnterKey: boolean = false
  ): void => {
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
        onUndo={handleUndo}
        onRedo={handleRedo}
      />

      <FormulaBar
        value={cells[selectedCell] || ""}
        onChange={(value: any) => handleCellChange(selectedCell, value)}
        selectedCell={selectedCell}
        formulaText={formulaText}
        onCellSelect={(cell: any) => {
          setSelectedCell(cell);
          setSelectionRange(null);
        }}
      />

      <div
        className="spreadsheet-container flex-1 overflow-auto relative"
        ref={gridRef}
      >
        {isResizing && resizeLinePosition && (
          <div
            className={`resize-reference-line ${resizeTypeRef.current}`}
            style={{
              [resizeTypeRef.current === "column"
                ? "left"
                : "top"]: `${resizeLinePosition}px`,
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
                    onMouseDown={(e) => handleResizeStart(e, "column", col)}
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
                    onMouseDown={(e) => handleResizeStart(e, "row", row)}
                  />
                </th>
                {Array.from({ length: numCols }, (_, col) => {
                  const cellName = getCellName(row, col);
                  const cellStyle = cellStyles[cellName] || {};
                  return (
                    <Cell
                      key={`${row}-${col}`}
                      name={cellName}
                      value={cells[cellName] || ""}
                      displayValue={getCellDisplayValue(cellName)}
                      onChange={(value: any) =>
                        handleCellChange(cellName, value)
                      }
                      onSelect={(e: any) => handleCellSelect(cellName, e)}
                      onMouseDown={(e: any) => handleMouseDown(cellName, e)}
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
};

export default SheetApp;
