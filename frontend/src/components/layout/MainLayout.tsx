import React from 'react';
import Menu from '../toolbar/Menu';
import Ribbon from '../toolbar/Ribbon';
import FormulaBar from '../toolbar/FormulaBar';
import ProjectExplorer from '../explorer/ProjectExplorer';
import DataTable from '../table/DataTable';
import NodeEditor from '../node-editor/NodeEditor';
import NodeProperties from '../node-editor/NodeProperties'; // Поместим в node-editor
import FunctionPanel from '../toolbar/FunctionPanel'; // Поместим в toolbar
import StatusBar from '../layout/StatusBar'; // Поместим в layout

const MainLayout: React.FC = () => {
  return (
    <div className="h-screen w-screen flex flex-col bg-gray-100 dark:bg-gray-900 text-gray-900 dark:text-gray-100 p-0 m-0 overflow-hidden">
      {/* Верхняя панель */}
      <div className="flex flex-col h-auto">
        <Menu />
        <Ribbon />
        <FormulaBar />
      </div>

      {/* Центральная область */}
      <div className="flex flex-1 overflow-hidden">
        {/* Левая панель (Обозреватель проекта) */}
        <div className="w-1/4 border-r border-gray-300 dark:border-gray-700 bg-white dark:bg-gray-800 flex flex-col">
          <ProjectExplorer />
        </div>

        {/* Правая панель (Таблица и Нодовый редактор сверху, Панель свойств снизу) */}
        <div className="flex-1 flex flex-col">
          {/* Верхняя часть (Таблица и Нодовый редактор) */}
          <div className="flex flex-1 overflow-hidden">
            {/* Таблица (2/3 ширины правой панели) */}
            <div className="flex-2 border-r border-gray-300 dark:border-gray-700 bg-white dark:bg-gray-800">
              <DataTable />
            </div>
            {/* Нодовый редактор (1/3 ширины правой панели) */}
            <div className="flex-1 bg-white dark:bg-gray-800">
              <NodeEditor />
            </div>
          </div>

          {/* Нижняя часть (Панель свойств) */}
          <div className="h-1/4 border-t border-gray-300 dark:border-gray-700 bg-white dark:bg-gray-800">
            <NodeProperties />
          </div>
        </div>
      </div>

      {/* Нижняя панель (Функции и Состояние) */}
      <div className="flex h-10 border-t border-gray-300 dark:border-gray-700 bg-gray-200 dark:bg-gray-700">
        <div className="w-1/3 border-r border-gray-300 dark:border-gray-600">
          <FunctionPanel />
        </div>
        <div className="flex-1">
          <StatusBar />
        </div>
      </div>
    </div>
  );
};

export default MainLayout;
