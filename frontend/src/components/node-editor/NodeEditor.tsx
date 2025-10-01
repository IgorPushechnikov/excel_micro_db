// frontend/src/components/node-editor/NodeEditor.tsx

import React from 'react';

const NodeEditor: React.FC = () => {
  return (
    <div className="h-full w-full flex flex-col">
      <div className="text-xs font-medium mb-2">Нодовый редактор</div>
      <div className="flex-1 bg-gray-200 dark:bg-gray-800 border border-dashed border-gray-400 dark:border-gray-500 rounded flex items-center justify-center">
        <p className="text-gray-500 dark:text-gray-400 text-sm">Здесь будет сцена для узлов</p>
      </div>
    </div>
  );
};

export default NodeEditor;