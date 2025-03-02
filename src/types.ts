export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  fontSize?: string;
  color?: string;
  fontFamily?: string;
  strikethrough?: boolean;
  textAlign?: 'left' | 'center' | 'right';
}

export interface CellData {
  value: string;
  formula?: string;
}

export interface SelectionRange {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export interface CellPosition {
  row: number;
  col: number;
}

export interface ResizeLinePosition {
  position: number;
  type: 'row' | 'column';
  index: number;
}

export type MathFunction = (values: number[]) => number;

export interface DataFunctionParams {
  value: string | string[];
  find?: string;
  replace?: string;
}

export type DataFunctionResult = string | Array<string | string[]>;

export type DataFunction = (params: DataFunctionParams) => DataFunctionResult;

export interface MathFunctions {
  [key: string]: MathFunction;
}

export interface DataFunctions {
  [key: string]: DataFunction;
}

export interface CellDimensions {
  [key: string]: number;
}

export interface FormulaBarProps {
  value: string;
  onChange: (cellName: string, value: string) => void;
  selectedCell: string;
  formulaText: string | null;
  onCellSelect: (cellName: string) => void;
}

export interface FunctionTesterProps {
  onClose: () => void;
  mathFunctions: MathFunctions;
  dataFunctions: DataFunctions;
  onSelect: (formula: string) => void;
}

export interface ContextMenuProps {
  x: number;
  y: number;
  onClose: () => void;
  onAddRow: () => void;
  onDeleteRow: () => void;
  onAddColumn: () => void;
  onDeleteColumn: () => void;
}

export interface ToolbarProps {
  onBold: () => void;
  onItalic: () => void;
  onFontSize: (size: string) => void;
  onColor: (color: string) => void;
  onFontFamily: (font: string) => void;
  onFormulaSelect: (formula: string) => void;
  selectedCell: string;
  cellStyles: Record<string, CellStyle>;
  currentFont: string;
  onSave: () => void;
  onLoad: () => void;
  onAlignment: (alignment: 'left' | 'center' | 'right') => void;
  onOpenFunctionTester: () => void;
}

export interface Formula {
  label: string;
  example: string;
}

export interface CellProps {
  value: string;
  displayValue: string;
  onChange: (value: string) => void;
  onSelect: (e: React.MouseEvent) => void;
  onMouseDown?: (e: React.MouseEvent) => void;
  onMouseMove?: (name: string) => void;
  isSelected: string | null;
  name: string;
  styles?: CellStyle;
  onDragStart?: (name: string) => void;
  onDragEnd?: () => void;
  onFillDragStart?: (name: string) => void;
  onFillDragMove?: (e: React.MouseEvent) => void;
  onFillDragEnd?: () => void;
  onAddRow: () => void;
  onDeleteRow: () => void;
  onAddColumn: () => void;
  onDeleteColumn: () => void;
  onTabNavigation: (name: string, shiftKey: boolean, isEnter?: boolean) => void;
} 