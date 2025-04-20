import React, { useState, useRef, useEffect, useCallback } from "react";
import ContextMenu from "./ContextMenu";

interface CellProps {
  value: string;
  displayValue: string;
  onChange: (value: string) => void;
  onSelect: (e: React.MouseEvent) => void;
  onMouseDown: (e: React.MouseEvent) => void;
  onMouseMove: (name: string) => void;
  isSelected: string;
  name: string;
  styles?: {
    fontWeight?: string;
    fontStyle?: string;
    fontSize?: string;
    color?: string;
    fontFamily?: string;
    textDecoration?: string;
    textAlign?: string;
  };
  onDragStart?: (name: string) => void;
  onDragEnd?: () => void;
  onFillDragStart?: (name: string) => void;
  onFillDragMove?: (e: React.MouseEvent) => void;
  onFillDragEnd?: () => void;
  onAddRow?: () => void;
  onDeleteRow?: () => void;
  onAddColumn?: () => void;
  onDeleteColumn?: () => void;
  onTabNavigation: (
    name: string,
    isShiftKey: boolean,
    isEnterKey?: boolean
  ) => void;
}

interface ContextMenuPosition {
  x: number;
  y: number;
}

const Cell: React.FC<CellProps> = ({
  value,
  displayValue,
  onChange,
  onSelect,
  onMouseDown,
  onMouseMove,
  isSelected,
  name,
  styles = {},
  onDragStart,
  onDragEnd,
  onFillDragStart,
  onFillDragMove,
  onFillDragEnd,
  onAddRow,
  onDeleteRow,
  onAddColumn,
  onDeleteColumn,
  onTabNavigation,
}) => {
  const [isEditing, setIsEditing] = useState<boolean>(false);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [isFillDragging, setIsFillDragging] = useState<boolean>(false);
  const [contextMenu, setContextMenu] = useState<ContextMenuPosition | null>(
    null
  );
  const cellRef = useRef<HTMLTableCellElement | null>(null);
  const fillHandleRef = useRef<HTMLDivElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);

  // Add keyboard event listener for the document
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Ignore if focus is on an input, select, or textarea element
      if (
        e.target instanceof Element &&
        (e.target.tagName === "INPUT" ||
          e.target.tagName === "SELECT" ||
          e.target.tagName === "TEXTAREA" ||
          e.target.closest(".toolbar") ||
          e.target.closest(".formula-bar"))
      ) {
        return;
      }

      // Only handle keypress if this cell is primary selected and not already editing
      if (isSelected?.includes("primary") && !isEditing) {
        // Handle tab navigation
        if (e.key === "Tab") {
          e.preventDefault();
          onTabNavigation(name, e.shiftKey);
          return;
        }

        // Handle Enter key for navigation
        if (e.key === "Enter") {
          e.preventDefault();
          // Move to the cell below
          const match = name.match(/([A-Z]+)(\d+)/);
          if (match) {
            const [_, col, row] = match;
            const nextCellName = `${col}${parseInt(row) + 1}`;
            onTabNavigation(nextCellName, false, true);
          }
          return;
        }

        // If it's a printable character or special keys like backspace/delete
        if (e.key.length === 1 || e.key === "Backspace" || e.key === "Delete") {
          setIsEditing(true);

          // For printable characters, we want to start fresh with just that character
          if (e.key.length === 1) {
            // Start with the character that was pressed
            onChange(e.key);
          } else {
            // For backspace/delete, we want to start with empty value
            onChange("");
          }

          e.preventDefault();
        }
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [isSelected, isEditing, onChange, name, onTabNavigation]);

  useEffect(() => {
    if (isSelected && isEditing && inputRef.current) {
      inputRef.current.focus();

      // Place cursor at the end of the text
      if (inputRef.current.value.length > 0) {
        inputRef.current.setSelectionRange(
          inputRef.current.value.length,
          inputRef.current.value.length
        );
      }
    }
  }, [isSelected, isEditing]);

  const handleDoubleClick = (): void => {
    setIsEditing(true);
  };

  const handleBlur = (): void => {
    setIsEditing(false);
  };

  const handleContextMenu = (e: React.MouseEvent): void => {
    e.preventDefault();
    setContextMenu({
      x: e.clientX,
      y: e.clientY,
    });
  };

  const handleCellMouseDown = useCallback(
    (e: React.MouseEvent): void => {
      if (e.button === 0) {
        // Left click only
        onSelect(e);
        if (!isEditing) {
          setIsDragging(true);
          onDragStart?.(name);
        }
      }
    },
    [isEditing, name, onSelect, onDragStart]
  );

  const handleCellMouseMove = useCallback(
    (e: React.MouseEvent | MouseEvent): void => {
      if (isDragging) {
        e.preventDefault();
        onMouseMove?.(name);
      } else if (isFillDragging && "clientX" in e) {
        e.preventDefault();
        onFillDragMove?.(e as React.MouseEvent);
      }
    },
    [isDragging, isFillDragging, name, onMouseMove, onFillDragMove]
  );

  const handleCellMouseUp = useCallback((): void => {
    if (isDragging) {
      setIsDragging(false);
      onDragEnd?.();
    }
    if (isFillDragging) {
      setIsFillDragging(false);
      onFillDragEnd?.();
    }
  }, [isDragging, isFillDragging, onDragEnd, onFillDragEnd]);

  const handleFillHandleMouseDown = useCallback(
    (e: React.MouseEvent): void => {
      e.preventDefault();
      e.stopPropagation();
      setIsFillDragging(true);
      onFillDragStart?.(name);
    },
    [name, onFillDragStart]
  );

  // Handle tab key navigation for input element
  const handleInputKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === "Tab") {
      e.preventDefault(); // Prevent default tab behavior
      setIsEditing(false);
      onTabNavigation(name, e.shiftKey);
    } else if (e.key === "Enter") {
      e.preventDefault();
      setIsEditing(false);
      // Move to the cell below
      const match = name.match(/([A-Z]+)(\d+)/);
      if (match) {
        const [_, col, row] = match;
        const nextCellName = `${col}${parseInt(row) + 1}`;
        onTabNavigation(nextCellName, false, true);
      }
    }
  };

  useEffect(() => {
    if (isDragging || isFillDragging) {
      document.addEventListener("mousemove", handleCellMouseMove);
      document.addEventListener("mouseup", handleCellMouseUp);

      return () => {
        document.removeEventListener("mousemove", handleCellMouseMove);
        document.removeEventListener("mouseup", handleCellMouseUp);
      };
    }
  }, [isDragging, isFillDragging, handleCellMouseMove, handleCellMouseUp]);

  const cellStyle: React.CSSProperties = {
    fontWeight: styles.fontWeight || "normal",
    fontStyle: styles.fontStyle || "normal",
    fontSize: styles.fontSize || "14px",
    color: styles.color || "#000000",
    fontFamily: styles.fontFamily || "Arial",
    textDecoration: styles.textDecoration || "none",
    // textAlign: styles.textAlign || "left",
    position: "relative",
  };

  return (
    <td
      ref={cellRef}
      className={`cell border border-gray-300 px-2 py-1 ${isSelected} ${
        isDragging ? "cell-dragging" : ""
      }`}
      style={cellStyle}
      data-cell-name={name}
      onMouseDown={handleCellMouseDown}
      onMouseMove={(e) => handleCellMouseMove(e)}
      onDoubleClick={handleDoubleClick}
      onContextMenu={handleContextMenu}
      draggable={false}
      tabIndex={isSelected?.includes("primary") ? 0 : -1}
    >
      {isEditing ? (
        <input
          ref={inputRef}
          type="text"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          onBlur={handleBlur}
          onKeyDown={handleInputKeyDown}
          className="w-full h-full outline-none"
          style={
            {
              // textAlign: styles.textAlign || "left",
            }
          }
          autoFocus
        />
      ) : (
        <div
          className="w-full h-full min-h-[20px]"
          style={
            {
              // textAlign: styles.textAlign || "left",
            }
          }
        >
          {displayValue}
        </div>
      )}
      {isSelected?.includes("primary") && !isEditing && (
        <>
          <div
            ref={fillHandleRef}
            className="fill-handle"
            onMouseDown={handleFillHandleMouseDown}
          />
          {isFillDragging && (
            <div
              className="autofill-cover"
              style={{
                top: fillHandleRef.current?.offsetTop || 0,
                left: fillHandleRef.current?.offsetLeft || 0,
              }}
            />
          )}
        </>
      )}
      {contextMenu && (
        <ContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          onClose={() => setContextMenu(null)}
          onAddRow={onAddRow}
          onDeleteRow={onDeleteRow}
          onAddColumn={onAddColumn}
          onDeleteColumn={onDeleteColumn}
        />
      )}
    </td>
  );
};

export default Cell;
