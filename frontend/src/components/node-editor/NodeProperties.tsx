// frontend/src/components/node-editor/NodeProperties.tsx

import React from 'react';

const NodeProperties: React.FC = () => {
  return (
    <div className="h-full w-full">
      <div className="text-xs font-medium mb-1">Панель свойств</div>
      <div className="text-xs">
        <p><span className="font-medium">Имя:</span> quarterly_summary</p>
        <p><span className="font-medium">Тип:</span> python_script</p>
        <p><span className="font-medium">Выход:</span> F2</p>
      </div>
    </div>
  );
};

export default NodeProperties;