// Базовый JavaScript для Excel Micro DB GUI - Простой нодовый редактор и ag-Grid

console.log('GUI Flask приложение загружено');

// --- ag-Grid ---

let gridApi;
let gridColumnApi;

function initAgGrid(nodeData = []) {
  const gridDiv = document.getElementById('ag-grid-container');
  if (!gridDiv) {
    console.error('Контейнер #ag-grid-container не найден');
    return;
  }

  // Определяем колонки
  const columnDefs = [
    { field: 'id', headerName: 'ID', sortable: true, filter: true },
    { field: 'name', headerName: 'Название', sortable: true, filter: true },
    { field: 'value', headerName: 'Значение', sortable: true, filter: true, type: 'numberColumn' },
  ];

  // Опции для ag-Grid
  const gridOptions = {
    defaultColDef: {
      resizable: true,
      minWidth: 100,
    },
    columnDefs: columnDefs,
    rowSelection: 'single',
    animateRows: true,
    enableCellTextSelection: true, // Позволяет выделять текст в ячейках
    onGridReady: (params) => {
      gridApi = params.api;
      gridColumnApi = params.columnApi;
      // Устанавливаем данные
      gridApi.setRowData(nodeData);
      console.log('ag-Grid инициализирован');
    }
  };

  // Инициализируем ag-Grid
  new agGrid.Grid(gridDiv, gridOptions);
}

// --- Функции взаимодействия с Flask ---

let currentProjectId = null;
let currentSheetName = null; // Для простоты будем работать с первым листом

function setStatus(message, isError = false) {
  const statusElement = document.getElementById('statusMessage');
  if (statusElement) {
    statusElement.textContent = message;
    statusElement.style.color = isError ? 'red' : 'green';
  }
}

async function uploadFile() {
  const fileInput = document.getElementById('fileInput');
  const file = fileInput.files[0];

  if (!file) {
    alert('Пожалуйста, выберите файл.');
    return;
  }

  const formData = new FormData();
  formData.append('file', file);

  try {
    setStatus('Загрузка файла...', false);
    const response = await fetch('/upload', {
      method: 'POST',
      body: formData
    });

    const data = await response.json();

    if (response.ok) {
      currentProjectId = data.project_id;
      setStatus(`Файл загружен: ${data.filename}`, false);
      document.getElementById('analyzeBtn').disabled = false;
      console.log('Файл загружен, project_id:', currentProjectId);
    } else {
      throw new Error(data.error || 'Неизвестная ошибка при загрузке');
    }
  } catch (error) {
    console.error('Ошибка при загрузке файла:', error);
    setStatus(`Ошибка загрузки: ${error.message}`, true);
  }
}

async function analyzeFile() {
  if (!currentProjectId) {
    alert('Нет активного проекта. Загрузите файл сначала.');
    return;
  }

  try {
    setStatus('Анализ файла...', false);
    const response = await fetch('/analyze', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({}) // Пока не передаем параметры
    });

    const data = await response.json();

    if (response.ok) {
      setStatus('Анализ завершен!', false);
      console.log('Анализ завершен:', data.message);
      // После анализа загружаем данные первого листа
      await loadSheetsAndData();
    } else {
      throw new Error(data.error || 'Неизвестная ошибка при анализе');
    }
  } catch (error) {
    console.error('Ошибка при анализе файла:', error);
    setStatus(`Ошибка анализа: ${error.message}`, true);
  }
}

async function loadSheetsAndData() {
  if (!currentProjectId) {
    console.warn('Нет активного проекта для загрузки листов.');
    return;
  }

  try {
    // 1. Получаем список листов
    const sheetsResponse = await fetch('/sheets');
    const sheetsData = await sheetsResponse.json();

    if (!sheetsResponse.ok) {
      throw new Error(sheetsData.error || 'Ошибка при получении списка листов');
    }

    const sheetNames = sheetsData.sheets;
    if (sheetNames.length === 0) {
      console.warn('В проекте нет листов.');
      return;
    }

    // 2. Берем первый лист
    currentSheetName = sheetNames[0];
    console.log(`Загрузка данных для листа: ${currentSheetName}`);

    // 3. Получаем данные листа
    const dataResponse = await fetch(`/sheet_data/${encodeURIComponent(currentSheetName)}`);
    const gridData = await dataResponse.json();

    if (!dataResponse.ok) {
      throw new Error(`Ошибка при получении данных листа ${currentSheetName}`);
    }

    // 4. Обновляем ag-Grid
    if (gridApi) {
      // Определяем columnDefs на основе первой строки данных
      let columnDefs = [];
      if (gridData.length > 0) {
        const firstRow = gridData[0];
        columnDefs = Object.keys(firstRow).map(key => ({
          field: key,
          headerName: key,
          sortable: true,
          filter: true,
          resizable: true,
          minWidth: 100
        }));
      }
      
      gridColumnApi.setColumnDefs(columnDefs);
      gridApi.setRowData(gridData);
      console.log(`Данные листа '${currentSheetName}' загружены в ag-Grid.`);
      setStatus(`Данные листа '${currentSheetName}' загружены.`, false);
    } else {
      // Если gridApi еще не инициализирован, инициализируем с новыми данными
      // Определяем columnDefs на основе первой строки данных
      let columnDefs = [];
      if (gridData.length > 0) {
        const firstRow = gridData[0];
        columnDefs = Object.keys(firstRow).map(key => ({
          field: key,
          headerName: key,
          sortable: true,
          filter: true,
          resizable: true,
          minWidth: 100
        }));
      }
      
      const gridDiv = document.getElementById('ag-grid-container');
      if (gridDiv) {
        // Уничтожаем предыдущую Grid, если она была
        if (gridApi) {
          gridApi.destroy();
          gridApi = null;
          gridColumnApi = null;
        }
        
        const gridOptions = {
          defaultColDef: {
            resizable: true,
            minWidth: 100,
          },
          columnDefs: columnDefs,
          rowData: gridData,
          rowSelection: 'single',
          animateRows: true,
          enableCellTextSelection: true,
          onGridReady: (params) => {
            gridApi = params.api;
            gridColumnApi = params.columnApi;
            console.log('ag-Grid инициализирован с данными листа');
          }
        };
        
        new agGrid.Grid(gridDiv, gridOptions);
        console.log(`ag-Grid инициализирован с данными листа '${currentSheetName}'.`);
        setStatus(`ag-Grid инициализирован с данными листа '${currentSheetName}'.`, false);
      }
    }

  } catch (error) {
    console.error('Ошибка при загрузке данных листов:', error);
    setStatus(`Ошибка загрузки данных: ${error.message}`, true);
  }
}

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
    // Инициализируем ag-Grid пустыми данными
    initAgGrid();
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

    // Добавим обработчик клика на узел для выбора
    this.container.addEventListener('click', (e) => {
      if (e.target.classList.contains('node')) {
        const nodeId = e.target.id;
        this.selectNode(nodeId);
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

    // --- Добавляем обработчики для новых кнопок ---
    document.getElementById('uploadBtn').addEventListener('click', uploadFile);
    document.getElementById('analyzeBtn').addEventListener('click', analyzeFile);
    // ----------------------------------------------------
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
      y: y,
      // Добавим поле для данных узла
      data: [
        { id: 1, name: `Данные ${nodeId}_1`, value: Math.random() * 100 },
        { id: 2, name: `Данные ${nodeId}_2`, value: Math.random() * 100 },
        { id: 3, name: `Данные ${nodeId}_3`, value: Math.random() * 100 },
      ]
    };

    this.nodes.push(nodeData);
    console.log('Добавлен узел:', nodeData);
  }

  selectNode(nodeId) {
    const node = this.nodes.find(n => n.id === nodeId);
    if (node) {
      this.selectedNode = node;
      console.log('Выбран узел:', node.id);
      // Обновляем ag-Grid данными выбранного узла
      if (gridApi) {
        gridApi.setRowData(node.data);
        // Опционально: скроллим к первой строке
        gridApi.ensureIndexVisible(0);
      } else {
        // Если gridApi ещё не готов, инициализируем с этими данными
        initAgGrid(node.data);
      }
    }
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
