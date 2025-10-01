// frontend/src/components/table/DataTable.tsx

import React from 'react';

const DataTable: React.FC = () => {
  // Простая демонстрация таблицы с заголовками и данными
  const headers = ['A', 'B', 'C', 'D'];
  const data = [
    ['Pr', 'Q1', 'Q2', 'Q3'],
    ['A', '100', '150', '200'],
    ['', '', '', ''], // Пустая строка для демонстрации
    ['', '', '', ''],
  ];

  return (
    <div className="overflow-auto h-full w-full">
      <table className="min-w-full border-collapse">
        <thead>
          <tr>
            <th className="border border-gray-300 p-1 w-16 bg-gray-100 dark:bg-gray-700 text-xs"></th> {/* Пустая ячейка для номеров строк */}
            {headers.map((header, index) => (
              <th 
                key={index} 
                className="border border-gray-300 p-1 w-24 bg-gray-200 dark:bg-gray-600 text-xs font-medium text-center"
              >
                {header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map((row, rowIndex) => (
            <tr key={rowIndex}>
              <td className="border border-gray-300 p-1 text-xs text-center bg-gray-100 dark:bg-gray-700">{rowIndex + 1}</td> {/* Номер строки */}
              {row.map((cell, cellIndex) => (
                <td 
                  key={cellIndex} 
                  className="border border-gray-300 p-1 text-xs"
                >
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default DataTable;