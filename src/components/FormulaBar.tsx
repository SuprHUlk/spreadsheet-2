import React, { useRef, useEffect, useState, ChangeEvent } from 'react';
import { FormulaBarProps } from '../types';

const FormulaBar: React.FC<FormulaBarProps> = ({ 
  value, 
  onChange, 
  selectedCell, 
  formulaText, 
  onCellSelect 
}) => {
  const inputRef = useRef<HTMLInputElement>(null);
  const cellInputRef = useRef<HTMLInputElement>(null);
  const [cellInputValue, setCellInputValue] = useState<string>(selectedCell);

  useEffect(() => {
    setCellInputValue(selectedCell);
  }, [selectedCell]);

  useEffect(() => {
    if (formulaText && inputRef.current) {
      inputRef.current.value = formulaText;
      inputRef.current.focus();
      // Place cursor between the parentheses
      const cursorPos = formulaText.indexOf('(') + 1;
      inputRef.current.setSelectionRange(cursorPos, cursorPos);
    }
  }, [formulaText]);

  const handleCellInputChange = (e: ChangeEvent<HTMLInputElement>): void => {
    const newValue = e.target.value.toUpperCase();
    setCellInputValue(newValue);
    
    // Only update selection if it's a valid cell reference
    if (/^[A-Z]+[1-9][0-9]*$/.test(newValue)) {
      onCellSelect(newValue);
    }
  };

  const handleCellInputBlur = (): void => {
    // Reset to selected cell if invalid input
    if (!/^[A-Z]+[1-9][0-9]*$/.test(cellInputValue)) {
      setCellInputValue(selectedCell);
    }
  };

  return (
    <div className="formula-bar flex items-center gap-1 py-1 px-2 border-b border-gray-300">
      <input
        ref={cellInputRef}
        type="text"
        className="cell-reference bg-gray-100 px-1 py-0.5 rounded text-center font-mono text-xs min-h-[20px] flex-none"
        value={cellInputValue}
        onChange={handleCellInputChange}
        onBlur={handleCellInputBlur}
        placeholder="A1"
      />
      <div className="formula-input flex-1">
        <input
          ref={inputRef}
          type="text"
          className="w-full p-1 border border-gray-300 rounded text-sm"
          value={value}
          onChange={(e) => onChange(selectedCell, e.target.value)}
          placeholder="Enter value or formula (e.g., =SUM(A1:B2))"
        />
      </div>
      <div className="formula-help text-xs text-gray-500 flex-none">
        {value?.startsWith('=') ? 'Formula Mode' : 'Value Mode'}
      </div>
    </div>
  );
};

export default FormulaBar; 