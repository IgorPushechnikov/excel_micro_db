// frontend/src/components/toolbar/Menu.jsx

import React from 'react';

const Menu = ({ menuConfig, onAction }) => {
  const handleItemClick = (item) => {
    if (item.action && onAction) {
      onAction(item.action, item);
    }
  };

  if (!menuConfig || !menuConfig.menu) {
    return <div className="text-sm font-medium">Меню: не загружено</div>;
  }

  return (
    <div className="flex items-center text-sm font-medium">
      {menuConfig.menu.map((section) => (
        <div key={section.label} className="relative group px-2 py-1">
          <span className="hover:underline cursor-pointer">
            {section.label}
          </span>
          {/* Простое всплывающее меню (без подменю в этом примере) */}
          <div className="absolute left-0 mt-1 w-48 bg-white dark:bg-gray-800 border border-gray-300 dark:border-gray-600 rounded shadow-lg z-40 hidden group-hover:block">
            <ul className="py-1">
              {section.items && section.items.map((item) => (
                item.separator ? (
                  <li key={`separator-${item.id || Math.random()}`} className="border-t border-gray-200 dark:border-gray-700 my-1"></li>
                ) : (
                  <li
                    key={item.id}
                    className="px-4 py-1 hover:bg-gray-200 dark:hover:bg-gray-700 cursor-pointer text-sm"
                    onClick={() => handleItemClick(item)}
                  >
                    {item.label}
                  </li>
                )
              ))}
            </ul>
          </div>
        </div>
      ))}
    </div>
  );
};

export default Menu;