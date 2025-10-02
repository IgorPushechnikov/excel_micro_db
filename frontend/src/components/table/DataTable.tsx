import React from 'react';

const DataTable: React.FC = () => {
  return (
    <div className="h-full w-full p-1 bg-yellow-200 border-2 border-red-500"> {/* Жёлтый фон и красная рамка для видимости */}
      <div className="h-full w-full flex items-center justify-center bg-blue-100">
        <p className="text-lg font-semibold text-gray-800">Таблица (DataTable) - Заглушка</p>
      </div>
    </div>
  );
};

export default DataTable;
