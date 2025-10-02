// src/components/toolbar/FormulaBar.jsx
import React from 'react';

const FormulaBar = () => {
  return (
    <div className="h-8 flex items-center px-2 bg-gray-200 dark:bg-gray-700 border-b border-gray-300 dark:border-gray-600">
      <span>Строка формул: </span>
      <input
        type="text"
        className="flex-1 ml-2 px-2 py-1 border border-gray-400 dark:border-gray-500 rounded bg-white dark:bg-gray-600 text-gray-900 dark:text-gray-100"
        value='=NODE("sales_validator") + SUM(B2:E2)'
        readOnly // Пока readOnly
        aria-label="Строка формул" // Для доступности
      />
    </div>
  );
};

export default FormulaBar;
