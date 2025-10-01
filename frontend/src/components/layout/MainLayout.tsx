// frontend/src/components/layout/MainLayout.tsx

import React from 'react';
import Menu from '../toolbar/Menu'; // Пока не существует, будет создано
import Ribbon from '../toolbar/Ribbon'; // Пока не существует, будет создано
import FormulaBar from '../toolbar/FormulaBar'; // Пока не существует, будет создано
import ProjectExplorer from '../explorer/ProjectExplorer'; // Пока не существует, будет создано
import DataTable from '../table/DataTable'; // Пока не существует, будет создано
import NodeEditor from '../node-editor/NodeEditor'; // Пока не существует, будет создано
import NodeProperties from '../node-editor/NodeProperties'; // Пока не существует, будет создано
// import StatusBar from './StatusBar'; // Удалено, так как компонент не используется напрямую

const MainLayout: React.FC = () => {
  return (
    <div className="flex flex-col h-screen w-screen bg-gray-100 dark:bg-gray-800"> {/* Основной контейнер */}
      {/* Меню */}
      <header className="bg-gray-200 dark:bg-gray-700 p-1 border-b border-gray-300 dark:border-gray-600">
        <Menu />
      </header>

      {/* Лента */}
      <div className="bg-gray-300 dark:bg-gray-600 p-1 border-b border-gray-300 dark:border-gray-500">
        <Ribbon />
      </div>

      {/* Строка формул */}
      <div className="bg-gray-200 dark:bg-gray-700 p-1 border-b border-gray-300 dark:border-gray-600">
        <FormulaBar />
      </div>

      {/* Основная область с делением на панели */}
      <div className="flex flex-1 overflow-hidden"> {/* flex-1 означает, что этот div занимает всё оставшееся место */}
        {/* Обозреватель проекта */}
        <div className="w-1/4 bg-gray-50 dark:bg-gray-900 p-2 border-r border-gray-300 dark:border-gray-600">
          <ProjectExplorer />
        </div>

        {/* Центральная область: Таблица и Нодовый редактор */}
        <div className="flex-1 flex flex-col"> {/* Центральная панель, занимает 1/2 оставшегося места */}
          <div className="flex flex-1 overflow-hidden"> {/* Верхняя часть центральной панели */}
            {/* Таблица */}
            <div className="flex-1 bg-white dark:bg-gray-800 p-2 border-r border-gray-300 dark:border-gray-600">
              <DataTable />
            </div>
            {/* Нодовый редактор */}
            <div className="w-1/3 bg-gray-100 dark:bg-gray-700 p-2">
              <NodeEditor />
            </div>
          </div>
          {/* Панель свойств */}
          <div className="h-1/4 bg-gray-50 dark:bg-gray-800 p-2 border-t border-gray-300 dark:border-gray-600">
            <NodeProperties />
          </div>
        </div>
      </div>

      {/* Нижняя панель: Функции и Состояние */}
      <footer className="flex h-16 bg-gray-200 dark:bg-gray-700 p-2 border-t border-gray-300 dark:border-gray-600">
        <div className="w-1/3 bg-gray-100 dark:bg-gray-600 p-1 rounded mr-2">
          {/* Функции */}
          <div className="text-xs font-medium text-gray-700 dark:text-gray-300">Функции</div>
        </div>
        <div className="flex-1 bg-gray-100 dark:bg-gray-600 p-1 rounded">
          {/* Состояние */}
          <div className="text-xs font-medium text-gray-700 dark:text-gray-300">Состояние: Готово | Последнее сохранение: 10:30</div>
        </div>
      </footer>
    </div>
  );
};

export default MainLayout;