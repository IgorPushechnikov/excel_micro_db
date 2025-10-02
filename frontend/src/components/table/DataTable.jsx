// src/components/table/DataTable.jsx (упрощённая версия на JS)
import React from 'react';

const DataTable = () => {
  const headers = ['A', 'B', 'C', 'D'];
  const rows = [
    ['Pr', 'Q1', 'Q2', 'Q3'],
    [100, 150, 200, 250],
    [120, 180, 220, 270],
    [110, 160, 210, 260],
    [130, 190, 240, 290],
    [140, 200, 250, 300],
    [150, 210, 260, 310],
    [160, 220, 270, 320],
    [170, 230, 280, 330],
    [180, 240, 290, 340],
  ];

  return (
    <div className="h-full w-full p-1 bg-white dark:bg-gray-800 overflow-auto">
      <table className="min-w-full border-collapse">
        <thead>
          <tr>
            {headers.map((header, i) => (
              <th key={i} className="border border-gray-300 px-2 py-1 bg-gray-100 dark:bg-gray-700">{header}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, cellIndex) => (
                <td key={cellIndex} className="border border-gray-300 px-2 py-1">{cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default DataTable;
