// frontend/src/components/explorer/ProjectExplorer.tsx

import React from 'react';

const ProjectExplorer: React.FC = () => {
  return (
    <div className="text-xs">
      <div className="font-medium mb-1">Обозреватель проекта</div>
      <ul className="space-y-1">
        <li className="p-1 hover:bg-gray-200 dark:hover:bg-gray-700 cursor-pointer rounded">• Sales_Data</li>
        <li className="p-1 hover:bg-gray-200 dark:hover:bg-gray-700 cursor-pointer rounded">• Mixed_Data</li>
        <li className="p-1 hover:bg-gray-200 dark:hover:bg-gray-700 cursor-pointer rounded">• Summary</li>
        <li className="p-1 hover:bg-gray-200 dark:hover:bg-gray-700 cursor-pointer rounded">• Edge_Cases</li>
      </ul>
    </div>
  );
};

export default ProjectExplorer;