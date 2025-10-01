// frontend/src/components/toolbar/FormulaBar.tsx

import React from 'react';

const FormulaBar: React.FC = () => {
  return (
    <div className="flex items-center">
      <span className="text-xs font-medium mr-2">=</span>
      <input 
        type="text" 
        placeholder="NODE(\"sales_validator\") + SUM(B2:E2)" 
        className="flex-1 text-xs p-1 border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
        readOnly // Пока readOnly, как заглушка
      />
    </div>
  );
};

export default FormulaBar;