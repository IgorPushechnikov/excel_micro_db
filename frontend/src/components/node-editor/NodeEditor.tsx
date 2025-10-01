// frontend/src/components/node-editor/NodeEditor.tsx
import React, { useCallback } from 'react';
import ReactFlow, {
  ReactFlowProvider,
  addEdge,
  Background,
  Controls,
  ControlButton,
  MiniMap,
  Connection,
  Edge,
  Node,
  NodeTypes,
  useNodesState,
  useEdgesState,
} from 'reactflow';
import 'reactflow/dist/style.css'; // Стили React Flow

// Импортируем кастомные типы узлов как default export
import FormulaNode from './NodeTypes/FormulaNode';
import PythonNode from './NodeTypes/PythonNode';
import DefaultNode from './NodeTypes/DefaultNode';

// Интерфейс для данных узла (может быть расширен)
interface NodeData {
  label: string;
  type?: string; // Тип скрипта: python, js, formula и т.д.
  // Другие поля...
}

// Определяем доступные типы узлов
const nodeTypes: NodeTypes = {
  formulaNode: FormulaNode,
  pythonNode: PythonNode,
  default: DefaultNode,
  // ... другие типы
};

// Начальные узлы и соединения
const initialNodes: Node<NodeData>[] = [
  {
    id: '1',
    type: 'formulaNode', // Используем кастомный тип
    position: { x: 0, y: 0 },
    data: { label: 'sales_validator', type: 'python' }, // Пример данных
  },
  {
    id: '2',
    type: 'pythonNode', // Используем кастомный тип
    position: { x: 300, y: 100 },
    data: { label: 'quarterly_summary', type: 'python' },
  },
];

const initialEdges: Edge[] = [
  {
    id: 'e1-2',
    source: '1',
    target: '2',
    animated: true, // Пример анимированного соединения
    style: { stroke: '#1e1e1e', strokeWidth: 2 }, // Цвет соединения в темной теме
  },
];

// Компонент обертка для ReactFlowProvider
const NodeEditorFlow = () => {
  const [nodes, setNodes, onNodesChange] = useNodesState<NodeData>(initialNodes);
  const [edges, setEdges, onEdgesChange] = useEdgesState(initialEdges);

  const onConnect = useCallback(
    (params: Connection) => setEdges((eds) => addEdge(params, eds)),
    [setEdges]
  );

  // Обработчик добавления нового узла (пока просто для примера)
  const onAddNode = useCallback(() => {
    const newNode: Node<NodeData> = {
      id: `node_${Date.now()}`, // Простой ID
      type: 'default',
      position: { x: Math.random() * 400, y: Math.random() * 400 },
      data: { label: `New Node ${nodes.length + 1}` },
    };
    setNodes((nds) => nds.concat(newNode));
  }, [nodes.length, setNodes]);

  return (
    <ReactFlow
      nodes={nodes}
      edges={edges}
      onNodesChange={onNodesChange}
      onEdgesChange={onEdgesChange}
      onConnect={onConnect}
      nodeTypes={nodeTypes} // Регистрируем кастомные типы
      fitView // При загрузке приближаем вид к узлам
      attributionPosition="bottom-left" // Позиция атрибуции React Flow
      // Дополнительные опции можно добавить здесь
    >
      {/* Фоновая сетка */}
      <Background gap={16} size={1} color="#aaa" />
      {/* Управление (масштаб, сброс) */}
      <Controls>
        <ControlButton onClick={onAddNode}>+</ControlButton> {/* Пример кнопки добавления */}
      </Controls>
      {/* Мини-карта */}
      <MiniMap nodeColor={(n) => {
        // Определяем цвет узла на мини-карте в зависимости от типа
        switch (n.type) {
          case 'formulaNode':
            return '#E8F5E8'; // Светло-зелёный
          case 'pythonNode':
            return '#E8F0FE'; // Светло-синий
          default:
            return '#E0E0E0'; // Серый по умолчанию
        }
      }} />
    </ReactFlow>
  );
};

// Основной компонент NodeEditor
const NodeEditor: React.FC = () => {
  return (
    <div className="h-full w-full">
      <div className="text-xs font-medium mb-2">Нодовый редактор</div>
      <div className="h-[calc(100%-1.5rem)] w-full border border-gray-300 dark:border-gray-600 rounded"> {/* Основной контейнер для React Flow */}
        <ReactFlowProvider>
          <NodeEditorFlow />
        </ReactFlowProvider>
      </div>
    </div>
  );
};

export default NodeEditor;