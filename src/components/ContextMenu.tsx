import React, { useEffect, useRef } from "react";

interface ContextMenuProps {
  x: number;
  y: number;
  onClose: () => void;
  onAddRow?: () => void;
  onDeleteRow?: () => void;
  onAddColumn?: () => void;
  onDeleteColumn?: () => void;
}

const ContextMenu: React.FC<ContextMenuProps> = ({
  x,
  y,
  onClose,
  onAddRow,
  onDeleteRow,
  onAddColumn,
  onDeleteColumn,
}) => {
  const menuRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        onClose();
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [onClose]);

  return (
    <div
      ref={menuRef}
      className="context-menu"
      style={{
        top: y,
        left: x,
      }}
    >
      <div className="context-menu-group">
        <button
          className="context-menu-item"
          onClick={() => {
            onAddRow && onAddRow();
            onClose();
          }}
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
            <path d="M12 5v14M5 12h14" />
          </svg>
          Insert Row Below
        </button>
        <button
          className="context-menu-item"
          onClick={() => {
            onDeleteRow && onDeleteRow();
            onClose();
          }}
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
            <path d="M5 12h14" />
          </svg>
          Delete Row
        </button>
      </div>

      <div className="context-menu-separator" />

      <div className="context-menu-group">
        <button
          className="context-menu-item"
          onClick={() => {
            onAddColumn && onAddColumn();
            onClose();
          }}
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
            <path d="M12 5v14M5 12h14" />
          </svg>
          Insert Column Right
        </button>
        <button
          className="context-menu-item"
          onClick={() => {
            onDeleteColumn && onDeleteColumn();
            onClose();
          }}
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
            <path d="M5 12h14" />
          </svg>
          Delete Column
        </button>
      </div>
    </div>
  );
};

export default ContextMenu;
