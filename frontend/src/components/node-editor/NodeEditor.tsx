// frontend/src/components/node-editor/NodeEditor.tsx
import React, { useCallback, useState, useEffect, useRef } from 'react'; // Добавлен useRef
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
  NodeProps,
  NodeMouseHandler,
  ReactFlowInstance,
} from 'reactflow';
import { Plus } from 'lucide-react'; // Импортируем иконку
import 'reactflow/dist/style.css'; // Стили React Flow

// Импортируем кастомные типы узлов как default export
import FormulaNode from './NodeTypes/FormulaNode';
import PythonNode from './NodeTypes/PythonNode';
import DefaultNode from './NodeTypes/DefaultNode';

// Импортируем конфигурации контекстных меню
import { paneMenuConfig, baseNodeMenuConfig } from '../../config/context_menus';

// Импортируем универсальный компонент контекстного меню
import ContextMenu from '../context_menus/ContextMenu';

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

  // Состояния для контекстного меню
  const [paneContextMenu, setPaneContextMenu] = useState<{ position: { x: number; y: number }; items: any[] } | null>(null);
  const [nodeContextMenu, setNodeContextMenu] = useState<{ position: { x: number; y: number }; items: any[]; nodeId: string } | null>(null);

  // Ссылка на экземпляр ReactFlow для получения позиции
  const reactFlowInstance = useRef<ReactFlowInstance<NodeData> | null>(null);

  const onConnect = useCallback(
    (params: Connection) => setEdges((eds) => addEdge(params, eds)),
    [setEdges]
  );

  // Обработчик добавления нового узла (пока просто для примера)
  const onAddNode = useCallback((nodeType: string, position?: { x: number; y: number }) => {
    if (!reactFlowInstance.current) return;

    // Если позиция не указана, используем случайную
    const pos = position || { x: Math.random() * 400, y: Math.random() * 400 };
    const newNode: Node<NodeData> = {
      id: `node_${Date.now()}`, // Простой ID
      type: nodeType,
      position: pos,
      data: { label: `New ${nodeType}`, type: 'python' },
    };
    setNodes((nds) => nds.concat(newNode));
  }, [setNodes]);

  // Обработчик удаления узла
  const onDeleteNode = useCallback((nodeId: string) => {
    if (!reactFlowInstance.current) return;
    reactFlowInstance.current.deleteElements({ nodes: [{ id: nodeId }] });
  }, []);

  // Обработчик дублирования узла
  const onDuplicateNode = useCallback((nodeId: string) => {
    const nodeToDuplicate = nodes.find(n => n.id === nodeId);
    if (!nodeToDuplicate) return;

    const duplicatedNode: Node<NodeData> = {
      ...nodeToDuplicate,
      id: `node_${Date.now()}`, // Новый уникальный ID
      position: { x: nodeToDuplicate.position.x + 100, y: nodeToDuplicate.position.y + 100 }, // Смещение
      data: { ...nodeToDuplicate.data }, // Копируем данные
    };
    setNodes((nds) => nds.concat(duplicatedNode));
  }, [nodes, setNodes]);

  // Обработчик добавления узла после другого
  const onAddNodeAfter = useCallback((sourceNodeId: string, nodeType: string) => {
    const sourceNode = nodes.find(n => n.id === sourceNodeId);
    if (!sourceNode) return;

    const newNodePos = { x: sourceNode.position.x + 250, y: sourceNode.position.y };
    onAddNode(nodeType, newNodePos);

    // Найдём ID нового узла (последний добавленный)
    // const newNodeId = `node_${Date.now() - 1}`; // Грубое приближение, лучше использовать ID из onAddNode
    // Пока используем временное решение, предполагая, что onAddNode добавляет узел с ID, содержащим текущее время
    // Это не идеально, но работает для демонстрации
    const newNodes = nodes.concat(); // Создаём копию
    const newNode = newNodes[newNodes.length - 1]; // Берём последний добавленный
    if (newNode) {
      // Добавим соединение sourceNode -> newNode
      const newEdge: Edge = {
        id: `e${sourceNodeId}-${newNode.id}`,
        source: sourceNodeId,
        target: newNode.id,
        animated: true,
        style: { stroke: '#1e1e1e', strokeWidth: 2 },
      };
      setEdges((eds) => eds.concat(newEdge));
    }
  }, [nodes, onAddNode, setEdges]);

  // Обработчик действий из меню
  const handleMenuAction = useCallback((action: string, item: any) => {
    switch (action) {
      case 'addNode':
        if (paneContextMenu) {
          onAddNode(item.nodeType, reactFlowInstance.current?.project({ x: paneContextMenu.position.x, y: paneContextMenu.position.y }));
          setPaneContextMenu(null);
        } else if (nodeContextMenu) {
          // Если вызвано из контекстного меню узла (например, "Добавить после")
          if (item.action === 'addNodeAfter' && nodeContextMenu.nodeId) {
            onAddNodeAfter(nodeContextMenu.nodeId, item.nodeType);
          }
          setNodeContextMenu(null);
        }
        break;
      case 'deleteNode':
        if (nodeContextMenu) {
          onDeleteNode(nodeContextMenu.nodeId);
          setNodeContextMenu(null);
        }
        break;
      case 'duplicateNode':
        if (nodeContextMenu) {
          onDuplicateNode(nodeContextMenu.nodeId);
          setNodeContextMenu(null);
        }
        break;
      // ... другие действия
      default:
        console.log(`Неизвестное действие: ${action}`, item);
        if (paneContextMenu) setPaneContextMenu(null);
        if (nodeContextMenu) setNodeContextMenu(null);
    }
  }, [paneContextMenu, nodeContextMenu, onAddNode, onDeleteNode, onDuplicateNode, onAddNodeAfter]);

  // Обработчик контекстного меню на панели
  const onPaneContextMenu: NodeMouseHandler = useCallback((event) => {
    event.preventDefault();
    setPaneContextMenu({
      position: { x: event.clientX, y: event.clientY },
      items: paneMenuConfig.menu,
    });
    setNodeContextMenu(null); // Закрываем меню узла, если оно открыто
  }, []);

  // Обработчик контекстного меню на узле
  const onNodeContextMenu = useCallback((event: React.MouseEvent, node: Node<NodeData> /*, _nodeInstance: NodeInstance<NodeData>*/) => {
    event.preventDefault();
    setNodeContextMenu({
      position: { x: event.clientX, y: event.clientY },
      items: baseNodeMenuConfig.menu, // Используем базовую конфигурацию
      nodeId: node.id,
    });
    setPaneContextMenu(null); // Закрываем меню панели, если оно открыто
  }, []);

  // Закрытие меню при клике вне
  useEffect(() => {
    const handleClick = () => {
      setPaneContextMenu(null);
      setNodeContextMenu(null);
    };
    document.addEventListener('click', handleClick);
    return () => document.removeEventListener('click', handleClick);
  }, []);

  // Функция для установки экземпляра ReactFlow
  const onInit = (instance: ReactFlowInstance<NodeData>) => {
    reactFlowInstance.current = instance;
  };

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
      onPaneContextMenu={onPaneContextMenu} // Обработчик для сцены
      onNodeContextMenu={onNodeContextMenu} // Обработчик для узлов
      onInit={onInit} // Установка ссылки на экземпляр
    >
      {/* Фоновая сетка */}
      <Background gap={16} size={1} color="#aaa" />
      {/* Управление (масштаб, сброс, добавление узла) */}
      <Controls>
        {/* Кнопка добавления узла с иконкой */}
        <ControlButton onClick={() => onAddNode('default')} title="Добавить узел">
          <Plus size={16} />
        </ControlButton>
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
      {/* Отображение контекстного меню панели */}
      {paneContextMenu && (
        <ContextMenu
          items={paneContextMenu.items}
          position={paneContextMenu.position}
          onClose={() => setPaneContextMenu(null)}
          onAction={handleMenuAction}
        />
      )}
      {/* Отображение контекстного меню узла */}
      {nodeContextMenu && (
        <ContextMenu
          items={nodeContextMenu.items}
          position={nodeContextMenu.position}
          onClose={() => setNodeContextMenu(null)}
          onAction={handleMenuAction}
        />
      )}
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