import React from 'react';
import { Handle, Position } from 'reactflow';

interface DefaultNodeProps {
  data: {
    label: string;
  };
}

const DefaultNode: React.FC<DefaultNodeProps> = ({ data }) => {
  return (
    <div className="px-4 py-2 shadow-md rounded text-sm font-medium bg-white border-2 border-gray-300">
      <Handle type="target" position={Position.Left} />
      <div>{data.label}</div>
      <Handle type="source" position={Position.Right} />
    </div>
  );
};

export default DefaultNode;
