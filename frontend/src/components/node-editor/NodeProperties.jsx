import React from 'react';

const NodeProperties = () => {
  // Пример данных, которые могут приходить из NodeEditor
  const selectedNodeData = {
    name: 'Default Node',
    type: 'default',
    output: 'C3',
  };

  return (
    <div className="h-full w-full p-2 bg-white dark:bg-gray-800 flex flex-col">
      <h3 className="text-lg font-semibold mb-2">Свойства узла</h3>
      {selectedNodeData ? (
        <div>
          <p className="mb-1"><strong>Имя:</strong> {selectedNodeData.name}</p>
          <p className="mb-1"><strong>Тип:</strong> {selectedNodeData.type}</p>
          <p className="mb-1"><strong>Выход:</strong> {selectedNodeData.output}</p>
        </div>
      ) : (
        <p>Узел не выбран</p>
      )}
    </div>
  );
};

export default NodeProperties;
