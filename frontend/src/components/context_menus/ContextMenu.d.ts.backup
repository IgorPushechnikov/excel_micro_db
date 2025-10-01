import { ComponentType } from 'react';

// Уточним тип для элемента меню
interface MenuItem {
  id?: string;
  label?: string;
  action?: string;
  nodeType?: string;
  separator?: boolean;
  // Добавьте другие возможные поля из YAML
}

declare const ContextMenu: ComponentType<{  items: MenuItem[]; // Уточнённый тип  position: { x: number; y: number } | null;  onClose: () => void;  onAction: (action: string, item: MenuItem) => void; // Уточнённый тип}>;

export default ContextMenu;