// Базовый JavaScript для Excel Micro DB GUI - Простой нодовый редактор

console.log('GUI Flask приложение загружено');

// --- Простой нодовый редактор ---

class SimpleNodeEditor {
  constructor(containerId) {
    this.container = document.getElementById(containerId);
    if (!this.container) {
      console.error(`Контейнер #${containerId} не найден`);
      return;
    }

    this.nodes = [];
    this.connections = [];
    this.selectedNode = null;
    this.draggingNode = null;
    this.offsetX = 0;
    this.offsetY = 0;

    this.initCanvas();
    this.bindEvents();
  }

  initCanvas() {
    this.container.style.position = 'relative';
    this.container.style.overflow = 'auto'; // Позволим прокрутку при необходимости
    this.container.style.backgroundColor = '#f0f0f0'; // Фон для холста
    this.container.style.border = '1px solid #ccc';
  }

  bindEvents() {
    this.container.addEventListener('click', (e) => {
      if (e.target === this.container) {
        this.addNode(e.offsetX, e.offsetY);
      }
    });

    this.container.addEventListener('mousedown', (e) => {
      if (e.target.classList.contains('node')) {
        this.startDrag(e);
      }
    });

    document.addEventListener('mousemove', (e) => {
      if (this.draggingNode) {
        this.drag(e);
      }
    });

    document.addEventListener('mouseup', () => {
      if (this.draggingNode) {
        this.stopDrag();
      }
    });
  }

  addNode(x, y) {
    const nodeId = `node_${Date.now()}`;
    const nodeElement = document.createElement('div');
    nodeElement.id = nodeId;
    nodeElement.className = 'node';
    nodeElement.style.position = 'absolute';
    nodeElement.style.left = `${x}px`;
    nodeElement.style.top = `${y}px`;
    nodeElement.style.width = '120px';
    nodeElement.style.height = '60px';
    nodeElement.style.backgroundColor = '#fff';
    nodeElement.style.border = '1px solid #000';
    nodeElement.style.borderRadius = '4px';
    nodeElement.style.display = 'flex';
    nodeElement.style.alignItems = 'center';
    nodeElement.style.justifyContent = 'center';
    nodeElement.style.cursor = 'move';
    nodeElement.textContent = 'Новый узел';

    this.container.appendChild(nodeElement);

    const nodeData = {
      id: nodeId,
      element: nodeElement,
      x: x,
      y: y
    };

    this.nodes.push(nodeData);
    console.log('Добавлен узел:', nodeData);
  }

  startDrag(e) {
    this.draggingNode = this.nodes.find(node => node.element === e.target);
    if (this.draggingNode) {
      this.offsetX = e.offsetX;
      this.offsetY = e.offsetY;
      this.draggingNode.element.style.zIndex = 1000; // Поднять узел при перетаскивании
      console.log('Начало перетаскивания узла:', this.draggingNode.id);
    }
  }

  drag(e) {
    if (this.draggingNode) {
      // Рассчитываем новую позицию, учитывая прокрутку контейнера
      const rect = this.container.getBoundingClientRect();
      const x = e.clientX - rect.left - this.offsetX + this.container.scrollLeft;
      const y = e.clientY - rect.top - this.offsetY + this.container.scrollTop;

      this.draggingNode.x = x;
      this.draggingNode.y = y;
      this.draggingNode.element.style.left = `${x}px`;
      this.draggingNode.element.style.top = `${y}px`;

      // Тут можно обновлять позиции соединений, если они есть
      this.updateConnections();
    }
  }

  stopDrag() {
    if (this.draggingNode) {
      this.draggingNode.element.style.zIndex = 'auto';
      console.log('Окончание перетаскивания узла:', this.draggingNode.id);
      this.draggingNode = null;
    }
  }

  updateConnections() {
    // Пока пусто, но сюда можно добавить обновление SVG/Canvas линий
    console.log('Обновление соединений...');
  }
}

// Запускаем инициализацию после загрузки DOM
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    new SimpleNodeEditor('node-editor-container');
  });
} else {
  new SimpleNodeEditor('node-editor-container');
}

// Место для инициализации нодового редактора и другой логики
