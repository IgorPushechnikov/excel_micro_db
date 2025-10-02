// src/components/node-editor/NodeTypes/FormulaNode.jsx
import React from 'react';
import { Handle, Position } from 'reactflow';

const FormulaNode = ({ data }) => {
  return (
    <div className="px-4 py-2 shadow-md rounded text-sm font-medium bg-green-200 border-2 border-green-500">
      <Handle type="target" position={Position.Left} />
      <div>{data.label}</div>
      <Handle type="source" position={Position.Right} />
    </div>
  );
};

export default FormulaNode;
