import React from "react";
import { useAppSelector } from "../redux/hooks";

const FONT_FAMILIES = [
  "Arial",
  "Times New Roman",
  "Courier New",
  "Georgia",
  "Verdana",
  "Helvetica",
  "Tahoma",
  "Trebuchet MS",
  "Impact",
];

// Formula types
interface Formula {
  label: string;
  example: string;
}

// Reorganize formulas into categories
const MATH_FORMULAS: Formula[] = [
  { label: "SUM", example: "=SUM(:)" },
  { label: "AVERAGE", example: "=AVERAGE(:)" },
  { label: "COUNT", example: "=COUNT(:)" },
  { label: "MAX", example: "=MAX(:)" },
  { label: "MIN", example: "=MIN(:)" },
];

const DATA_FORMULAS: Formula[] = [
  { label: "TRIM", example: "=TRIM()" },
  { label: "UPPER", example: "=UPPER()" },
  { label: "LOWER", example: "=LOWER()" },
  // { label: "REMOVE_DUPLICATES", example: "=REMOVE_DUPLICATES(:)" },
  // {
  //   label: "FIND_AND_REPLACE",
  //   example: '=FIND_AND_REPLACE(:,"find","replace")',
  // },
];

interface CellStyle {
  fontWeight?: string;
  fontStyle?: string;
  fontSize?: string;
  color?: string;
  fontFamily?: string;
  textDecoration?: string;
  textAlign?: string;
}

interface ToolbarProps {
  onBold: () => void;
  onItalic: () => void;
  onFontSize: (size: string) => void;
  onColor: (color: string) => void;
  onFontFamily: (font: string) => void;
  onFormulaSelect: (formula: string) => void;
  selectedCell: string;
  cellStyles: CellStyle;
  currentFont: string;
  onSave: () => void;
  onLoad: () => void;
  onAlignment: (alignment: string) => void;
  onOpenFunctionTester: () => void;
  onAddRow: (index: number) => void;
  onDeleteRow: (index: number) => void;
  onAddColumn: (index: number) => void;
  onDeleteColumn: (index: number) => void;
  onUndo: () => void;
  onRedo: () => void;
}

const Toolbar: React.FC<ToolbarProps> = ({
  onBold,
  onItalic,
  onFontSize,
  onColor,
  onFontFamily,
  onFormulaSelect,
  selectedCell,
  cellStyles,
  currentFont,
  onSave,
  onLoad,
  onAlignment,
  onOpenFunctionTester,
  onAddRow,
  onDeleteRow,
  onAddColumn,
  onDeleteColumn,
  onUndo,
  onRedo,
}) => {
  // Get undo/redo stack info from Redux store
  const { undoStack, redoStack } = useAppSelector((state) => state.spreadsheet);
  // Get last saved info from Redux store
  const lastSaved = useAppSelector((state) => state.spreadsheet.lastSaved);

  // Format the last saved date for display
  const formatLastSaved = () => {
    if (!lastSaved) return "Never";
    return new Date(lastSaved).toLocaleString();
  };

  return (
    <div className="toolbar bg-white p-1 flex items-center gap-1 border-b border-gray-300">
      <div className="flex flex-row items-center gap-2 mr-4">
        <button
          className="toolbar-btn flex items-center gap-1 px-3"
          onClick={onSave}
          title="Save as Excel"
        >
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
            <polyline points="17 21 17 13 7 13 7 21" />
            <polyline points="7 3 7 8 15 8" />
          </svg>
        </button>
        <button
          className="toolbar-btn flex items-center gap-1 px-3"
          onClick={onLoad}
          title="Load Excel File"
        >
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
            <polyline points="17 8 12 3 7 8" />
            <line x1="12" y1="3" x2="12" y2="15" />
          </svg>
        </button>

        {/* New Undo/Redo buttons */}
        <button
          className={`toolbar-btn flex items-center gap-1 px-3 ${
            undoStack.length === 0 ? "opacity-50 cursor-not-allowed" : ""
          }`}
          onClick={onUndo}
          title="Undo (Ctrl+Z)"
          disabled={undoStack.length === 0}
        >
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M3 10h10a8 8 0 0 1 8 8v2M3 10l6 6M3 10l6-6" />
          </svg>
        </button>
        <button
          className={`toolbar-btn flex items-center gap-1 px-3 ${
            redoStack.length === 0 ? "opacity-50 cursor-not-allowed" : ""
          }`}
          onClick={onRedo}
          title="Redo (Ctrl+Y)"
          disabled={redoStack.length === 0}
        >
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M21 10h-10a8 8 0 0 0-8 8v2M21 10l-6 6M21 10l-6-6" />
          </svg>
        </button>
      </div>

      <div className="divider"></div>

      <select
        className="toolbar-select w-32 text-sm"
        value={currentFont}
        onChange={(e) => onFontFamily(e.target.value)}
      >
        {FONT_FAMILIES.map((font) => (
          <option key={font} value={font} style={{ fontFamily: font }}>
            {font}
          </option>
        ))}
      </select>

      <div className="divider"></div>

      <button
        className="toolbar-btn"
        title="Decrease font size"
        onClick={() => {
          const currentSize = parseInt(cellStyles?.fontSize || "10");
          onFontSize(`${Math.max(8, currentSize - 1)}px`);
        }}
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="20"
          height="20"
          viewBox="0 0 24 24"
        >
          <path d="M19 13H5v-2h14v2z" />
        </svg>
      </button>

      <input
        type="text"
        className="toolbar-input w-12 text-center"
        value={parseInt(cellStyles?.fontSize || "10")}
        onChange={(e) => {
          const value = e.target.value.replace(/[^\d]/g, "");
          if (
            value === "" ||
            (parseInt(value) >= 1 && parseInt(value) <= 400)
          ) {
            onFontSize(`${value}px`);
          }
        }}
        onBlur={(e) => {
          const value = e.target.value;
          if (!value || isNaN(parseInt(value))) {
            onFontSize("10px");
          } else {
            const size = Math.max(1, Math.min(400, parseInt(value)));
            onFontSize(`${size}px`);
          }
        }}
        maxLength={3}
      />

      <button
        className="toolbar-btn"
        title="Increase font size"
        onClick={() => {
          const currentSize = parseInt(cellStyles?.fontSize || "10");
          onFontSize(`${Math.min(400, currentSize + 1)}px`);
        }}
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="20"
          height="20"
          viewBox="0 0 24 24"
        >
          <path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z" />
        </svg>
      </button>

      <div className="divider"></div>

      <button
        className={`toolbar-btn ${
          cellStyles?.fontWeight === "bold" ? "bg-gray-200" : ""
        }`}
        onClick={onBold}
        title="Bold (Ctrl+B)"
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="20"
          height="20"
          viewBox="0 0 24 24"
        >
          <path d="M15.6 10.79c.97-.67 1.65-1.77 1.65-2.79 0-2.26-1.75-4-4-4H7v14h7.04c2.09 0 3.71-1.7 3.71-3.79 0-1.52-.86-2.82-2.15-3.42zM10 6.5h3c.83 0 1.5.67 1.5 1.5s-.67 1.5-1.5 1.5h-3v-3zm3.5 9H10v-3h3.5c.83 0 1.5.67 1.5 1.5s-.67 1.5-1.5 1.5z" />
        </svg>
      </button>

      <button
        className={`toolbar-btn ${
          cellStyles?.fontStyle === "italic" ? "bg-gray-200" : ""
        }`}
        onClick={onItalic}
        title="Italic (Ctrl+I)"
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="20"
          height="20"
          viewBox="0 0 24 24"
        >
          <path d="M10 4v3h2.21l-3.42 8H6v3h8v-3h-2.21l3.42-8H18V4z" />
        </svg>
      </button>

      <input
        type="color"
        className="toolbar-btn w-8 h-8 p-1"
        onChange={(e) => onColor(e.target.value)}
        value={cellStyles?.color || "#000000"}
        title="Text color"
      />

      <div className="divider"></div>

      <div className="flex items-center align-btns">
        <button
          className={`toolbar-btn ${
            cellStyles?.textAlign === "left" ? "bg-gray-200" : ""
          }`}
          onClick={() => onAlignment("left")}
          title="Align Left"
        >
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="20"
            height="20"
            viewBox="0 0 24 24"
          >
            <path d="M3 3h18v2H3zm0 8h12v2H3zm0 8h18v2H3zm0-4h12v2H3z" />
          </svg>
        </button>
        <button
          className={`toolbar-btn ${
            cellStyles?.textAlign === "center" ? "bg-gray-200" : ""
          }`}
          onClick={() => onAlignment("center")}
          title="Align Center"
        >
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="20"
            height="20"
            viewBox="0 0 24 24"
          >
            <path d="M3 3h18v2H3zm3 8h12v2H6zm-3 8h18v2H3zm3-4h12v2H6zm-3-8h18v2H3z" />
          </svg>
        </button>
        <button
          className={`toolbar-btn ${
            cellStyles?.textAlign === "right" ? "bg-gray-200" : ""
          }`}
          onClick={() => onAlignment("right")}
          title="Align Right"
        >
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="20"
            height="20"
            viewBox="0 0 24 24"
          >
            <path d="M3 3h18v2H3zm6 8h12v2H9zm-6 8h18v2H3zm6-4h12v2H9zm-6-8h18v2H3z" />
          </svg>
        </button>
      </div>

      <div className="divider"></div>

      <select
        className="toolbar-select w-40 text-sm font-mono"
        onChange={(e) => onFormulaSelect(e.target.value)}
        value=""
      >
        <option value="">Functions</option>
        <optgroup label="Mathematical Functions">
          {MATH_FORMULAS.map((formula) => (
            <option key={formula.label} value={formula.example}>
              {formula.label} {formula.example}
            </option>
          ))}
        </optgroup>
        <optgroup label="Data Quality Functions">
          {DATA_FORMULAS.map((formula) => (
            <option key={formula.label} value={formula.example}>
              {formula.label} {formula.example}
            </option>
          ))}
        </optgroup>
      </select>

      <button
        className="toolbar-btn px-3"
        onClick={onOpenFunctionTester}
        title="Test Functions"
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="16"
          height="16"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
        >
          <path d="M4 19h16M4 15h16M4 11h16M4 7h16" />
          <circle cx="9" cy="11" r="1" />
        </svg>
        Test
      </button>
    </div>
  );
};

export default Toolbar;
