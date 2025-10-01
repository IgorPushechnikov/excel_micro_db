import React from 'react';

const StatusBar: React.FC = () => {
  return (
    <div className="h-full flex items-center px-2 justify-end">
      <span>Готово | Последнее сохранение: 10:30</span>
    </div>
  );
};

export default StatusBar;
