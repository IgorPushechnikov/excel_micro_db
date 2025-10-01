// frontend/src/components/node-editor/NodeTypes/DefaultNode.tsx

import React from 'react';
import { Handle, Position, NodeProps, useReactFlow } from 'reactflow';
import { X } from 'lucide-react'; // Импортируем иконку
import 'reactflow/dist/style.css';

// Интерфейс для данных узла (может быть расширен)
interface NodeData {
  label: string;
  type?: string;
  // Другие поля...
}

const DefaultNode: React.FC<NodeProps<NodeData>> = ({ id, data }) => {
  const { setNodes } = useReactFlow<NodeData>(); // Получаем функцию для изменения узлов

  const handleDelete = () => {
    // Удаляем текущий узел по его ID
    setNodes((nds) => nds.filter((node) => node.id !== id));
  };

  return (
    <div className="px-4 py-2 shadow-md rounded-md bg-white dark:bg-gray-700 border-2 border-gray-300 dark:border-gray-500 relative"> {/* Добавлен relative */}
      <Handle type="target" position={Position.Top} className="w-2 h-2 bg-gray-500" />
      <div className="text-sm font-medium text-gray-700 dark:text-gray-200">
        {data.label}
      </div>
      {/* Кнопка удаления */}
      <button 
        onClick={handleDelete}
        className="absolute top-1 right-1 w-5 h-5 flex items-center justify-center text-gray-500 hover:text-red-500 hover:bg-gray-200 dark:hover:bg-gray-600 rounded-full"
        title="Удалить узел"
      >
        <X size={12} />
      </button>
      <Handle type="source" position={Position.Bottom} className="w-2 h-2 bg-gray-500" />
    </div>
  );
};

export default DefaultNode;