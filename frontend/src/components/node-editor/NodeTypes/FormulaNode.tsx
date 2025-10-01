// frontend/src/components/node-editor/NodeTypes/FormulaNode.tsx

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

const FormulaNode: React.FC<NodeProps<NodeData>> = ({ id, data }) => {
  const { setNodes } = useReactFlow<NodeData>(); // Получаем функцию для изменения узлов

  const handleDelete = () => {
    // Удаляем текущий узел по его ID
    setNodes((nds) => nds.filter((node) => node.id !== id));
  };

  return (
    <div className="px-4 py-2 shadow-md rounded-md bg-green-100 dark:bg-green-900 border-2 border-green-300 dark:border-green-700 relative"> {/* Добавлен relative */}
      <Handle type="target" position={Position.Left} className="w-2 h-2 bg-green-500" />
      <div className="text-sm font-medium text-green-800 dark:text-green-200">
        {data.label}
      </div>
      {/* Кнопка удаления */}
      <button 
        onClick={handleDelete}
        className="absolute top-1 right-1 w-5 h-5 flex items-center justify-center text-green-500 hover:text-red-500 hover:bg-green-200 dark:hover:bg-green-800 rounded-full"
        title="Удалить узел"
      >
        <X size={12} />
      </button>
      <Handle type="source" position={Position.Right} className="w-2 h-2 bg-green-500" />
    </div>
  );
};

export default FormulaNode;