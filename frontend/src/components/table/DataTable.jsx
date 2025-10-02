// excel_micro_db/frontend/src/components/table/DataTable.jsx
import React, { useMemo } from 'react';
import { AgGridReact } from '@ag-grid-community/react';
import { ModuleRegistry } from '@ag-grid-community/core';
import { ClientSideRowModelModule } from '@ag-grid-community/client-side-row-model';

// Регистрируем необходимый модуль
ModuleRegistry.registerModules([ClientSideRowModelModule]);

// Импортируем стили ag-Grid
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';

const DataTable = () => {
  // Определяем колонки с использованием useMemo
  const columnDefs = useMemo(() => [
    { field: 'A', headerName: 'A', editable: true, sortable: true, filter: true },
    { field: 'B', headerName: 'B', editable: true, sortable: true, filter: true },
    { field: 'C', headerName: 'C', editable: true, sortable: true, filter: true },
    { field: 'D', headerName: 'D', editable: true, sortable: true, filter: true },
  ], []);

  // Определяем данные
  const rowData = [
    { A: 'Pr', B: 'Q1', C: 'Q2', D: 'Q3' },
    { A: 'A', B: 100, C: 150, D: 200 },
    { A: 'B', B: 120, C: 180, D: 220 },
    { A: 'C', B: 110, C: 160, D: 210 },
    { A: 'D', B: 130, C: 190, D: 230 },
    { A: 'E', B: 140, C: 200, D: 240 },
    { A: 'F', B: 150, C: 210, D: 250 },
    { A: 'G', B: 160, C: 220, D: 260 },
    { A: 'H', B: 170, C: 230, D: 270 },
    { A: 'I', B: 180, C: 240, D: 280 },
  ];

  // Опции для ag-Grid
  const gridOptions = {
    rowSelection: 'multiple', // Явно указываем тип как 'multiple' | 'single'
    enableCellTextSelection: true,
    onGridReady: (_params) => { // Используем _ для неиспользуемого параметра
      console.log('ag-Grid ready'); // Для отладки
    },
    // Возможность добавить другие опции ag-Grid
  };

  return (
    <div className="h-full w-full"> {/* Внешний контейнер */}
      <div className="h-full w-full p-1"> {/* Внутренний контейнер с отступом */}
        <div
          id="myGrid"
          className="ag-theme-alpine-dark h-full w-full" // Контейнер для AgGridReact с темой и размерами
        >
          <AgGridReact
            rowData={rowData}
            columnDefs={columnDefs}
            gridOptions={gridOptions}
            // rowModelType больше не нужно указывать, но модуль зарегистрирован
          />
        </div>
      </div>
    </div>
  );
};

export default DataTable;
