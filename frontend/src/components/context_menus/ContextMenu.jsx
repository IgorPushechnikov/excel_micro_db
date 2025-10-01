// frontend/src/components/context_menus/ContextMenu.jsx

import React from 'react';

const ContextMenu = ({ items, position, onClose, onAction }) => {
  if (!position) return null;

  const handleItemClick = (item) => {
    if (item.action && onAction) {
      onAction(item.action, item);
    }
    onClose(); // Закрыть меню после клика
  };

  const handleContextMenu = (e) => {
    // Предотвращаем всплытие контекстного меню компонента до браузерского
    e.preventDefault();
    e.stopPropagation(); 
  };

  return (
    <div
      className="absolute bg-white dark:bg-gray-800 border border-gray-300 dark:border-gray-600 rounded shadow-lg z-50"
      style={{ left: position.x, top: position.y }}
      onContextMenu={handleContextMenu} // Предотвращаем браузерное контекстное меню на самом меню
    >
      <ul className="py-1">
        {items.map((item, index) => (
          item.separator ? (
            <li key={`separator-${index}`} className="border-t border-gray-200 dark:border-gray-700 my-1"></li>
          ) : (
            <li
              key={item.id || index}
              className="px-4 py-1 hover:bg-gray-200 dark:hover:bg-gray-700 cursor-pointer text-sm"
              onClick={() => handleItemClick(item)}
            >
              {item.label}
            </li>
          )
        ))}
      </ul>
    </div>
  );
};

export default ContextMenu;