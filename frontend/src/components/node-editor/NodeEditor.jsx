// src/components/node-editor/NodeEditor.jsx
import React, { useCallback, useState } from 'react';
import ReactFlow, {
  MiniMap,
  Controls,
  Background,
  useNodesState,
  useEdgesState,
  addEdge,
  Connection,
  Edge,
  Node,
  Panel,
} from 'reactflow';

import 'reactflow/dist/style.css';

// Импортируем кастомные узлы
import DefaultNode from './NodeTypes/DefaultNode';
import FormulaNode from './NodeTypes/FormulaNode';
import PythonNode from './NodeTypes/PythonNode';

// Определяем типы узлов
const nodeTypes = {
  defaultNode: DefaultNode,
  formulaNode: FormulaNode,
  pythonNode: PythonNode,
};

// Начальные узлы
const initialNodes = [
  { id: '1', type: 'formulaNode', position: { x: 0, y: 0 }, data: { label: 'sales_validator' } },
  { id: '2', type: 'pythonNode', position: { x: 200, y: 100 }, data: { label: 'quarterly_sum...' } },
];

// Начальные ребра
const initialEdges = [
  { id: 'e1-2', source: '1', target: '2', sourceHandle: 'source', targetHandle: 'target' },
];

const NodeEditor = () => {
  const [nodes, setNodes, onNodesChange] = useNodesState(initialNodes);
  const [edges, setEdges, onEdgesChange] = useEdgesState(initialEdges);

  // Обработчик соединений
  const onConnect = useCallback(
    (params) => setEdges((eds) => addEdge(params, eds)),
    [setEdges]
  );

  return (
    <div className="h-full w-full">
      <ReactFlow
        nodes={nodes}
        edges={edges}
        onNodesChange={onNodesChange}
        onEdgesChange={onEdgesChange}
        onConnect={onConnect}
        nodeTypes={nodeTypes}
        fitView
      >
        <MiniMap />
        <Controls />
        <Background gap={12} size={1} />
        <Panel position="top-right">
          <button className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-1 px-2 rounded text-xs">
            Добавить узел
          </button>
        </Panel>
      </ReactFlow>
    </div>
  );
};

export default NodeEditor;
