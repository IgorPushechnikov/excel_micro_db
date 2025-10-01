declare module "*.yaml" {
  const content: any; // Или уточните тип, если знаете структуру
  export default content;
}

declare module "*.yml" {
  const content: any; // Или уточните тип, если знаете структуру
  export default content;
}

// Импортируем и экспортируем конкретные конфигурации
import paneMenuConfig from './node_editor_pane.yaml';
import baseNodeMenuConfig from './node_types/base_node.yaml';

export {
  paneMenuConfig,
  baseNodeMenuConfig,
};