// frontend/src/components/table/DataTable.tsx
import React, { useState, useEffect } from 'react';
import { AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';

// Опционально: определение типа для строки данных
interface RowData {
  id: number;
  [key: string]: any; // Позволяет другим динамическим полям
}

const DataTable: React.FC = () => {
  // Состояния для колонок и данных
  const [columnDefs, setColumnDefs] = useState<any[]>([]);
  const [rowData, setRowData] = useState<RowData[]>([]);

  // Инициализация данных и колонок при монтировании компонента
  useEffect(() => {
    // Пример начальных колонок (A, B, C, D)
    const initialCols = Array.from({ length: 4 }, (_, i) => ({
      field: String.fromCharCode(65 + i), // 'A', 'B', 'C', 'D'
      editable: true,
      sortable: true,
      filter: true,
      resizable: true,
      minWidth: 100,
    }));

    // Пример начальных данных
    const initialRows = Array.from({ length: 10 }, (_, rowIndex) => {
      const row: RowData = { id: rowIndex + 1 };
      initialCols.forEach((col, colIndex) => {
        row[col.field] = rowIndex === 0 ? col.field : `Cell ${rowIndex + 1}${col.field}`;
      });
      return row;
    });

    setColumnDefs(initialCols);
    setRowData(initialRows);
  }, []);

  // Базовая конфигурация для ag-Grid
  const gridOptions = {
    defaultColDef: {
      editable: true,
      sortable: true,
      filter: true,
      resizable: true,
    },
    rowSelection: 'multiple', // Позволяет выделять несколько строк
    animateRows: true, // Анимация при добавлении/удалении строк
    domLayout: 'normal', // Важно для Tailwind, чтобы высота работала корректно
  };

  return (
    <div
      className="ag-theme-alpine h-full w-full border border-gray-300 dark:border-gray-600"
      style={{ height: '100%', width: '100%' }}
    >
      <AgGridReact
        rowData={rowData}
        columnDefs={columnDefs}
        gridOptions={gridOptions}
        // Дополнительные пропсы можно добавить здесь
        // например, onCellValueChanged для отслеживания изменений
      />
    </div>
  );
};

export default DataTable;