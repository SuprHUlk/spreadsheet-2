import { createSlice, PayloadAction } from "@reduxjs/toolkit";

// Types
export interface CellsState {
  [key: string]: string;
}

export interface CellStylesState {
  [key: string]: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontSize?: string;
    fontFamily?: string;
    textColor?: string;
    bgColor?: string;
    textDecoration?: string;
    textAlign?: "left" | "center" | "right";
  };
}

// Define the state structure
export interface SpreadsheetState {
  cells: CellsState;
  cellStyles: CellStylesState;
  // columnWidths: Record<number, number>;
  // rowHeights: Record<number, number>;
  undoStack: { cells: CellsState; cellStyles: CellStylesState }[];
  redoStack: { cells: CellsState; cellStyles: CellStylesState }[];
  lastSaved: Date | null;
}

// Initial state
const initialState: SpreadsheetState = {
  cells: {},
  cellStyles: {},
  // columnWidths: {},
  // rowHeights: {},
  undoStack: [],
  redoStack: [],
  lastSaved: null,
};

// Create the slice
const spreadsheetSlice = createSlice({
  name: "spreadsheet",
  initialState,
  reducers: {
    // Update a cell's content
    updateCell: (
      state,
      action: PayloadAction<{ cellName: string; value: string }>
    ) => {
      const { cellName, value } = action.payload;

      // Push to undo stack before making changes
      state.undoStack.push({
        cells: { ...state.cells },
        cellStyles: { ...state.cellStyles },
      });

      // Clear redo stack on new changes
      state.redoStack = [];

      // Update the cell
      state.cells[cellName] = value;
    },

    // Update a cell's style
    updateCellStyle: (
      state,
      action: PayloadAction<{
        cellName: string;
        styleProperty: string;
        value: any;
      }>
    ) => {
      const { cellName, styleProperty, value } = action.payload;

      // Push to undo stack before making changes
      state.undoStack.push({
        cells: { ...state.cells },
        cellStyles: { ...state.cellStyles },
      });

      // Clear redo stack
      state.redoStack = [];

      // Initialize cell style object if it doesn't exist
      if (!state.cellStyles[cellName]) {
        state.cellStyles[cellName] = {};
      }

      // Update the style property
      state.cellStyles[cellName] = {
        ...state.cellStyles[cellName],
        [styleProperty]: value,
      };
    },

    // Undo last action
    undo: (state) => {
      if (state.undoStack.length > 0) {
        // Save current state to redo stack
        state.redoStack.push({
          cells: { ...state.cells },
          cellStyles: { ...state.cellStyles },
        });

        // Pop and apply the last state from undo stack
        const lastState = state.undoStack.pop();
        if (lastState) {
          state.cells = lastState.cells;
          state.cellStyles = lastState.cellStyles;
        }
      }
    },

    // Redo last undone action
    redo: (state) => {
      if (state.redoStack.length > 0) {
        // Save current state to undo stack
        state.undoStack.push({
          cells: { ...state.cells },
          cellStyles: { ...state.cellStyles },
        });

        // Pop and apply the last state from redo stack
        const nextState = state.redoStack.pop();
        if (nextState) {
          state.cells = nextState.cells;
          state.cellStyles = nextState.cellStyles;
        }
      }
    },

    // Update last saved timestamp
    setLastSaved: (state) => {
      state.lastSaved = new Date();
    },
  },
});

// Export actions and reducer
export const { updateCell, updateCellStyle, undo, redo, setLastSaved } =
  spreadsheetSlice.actions;
export default spreadsheetSlice.reducer;
