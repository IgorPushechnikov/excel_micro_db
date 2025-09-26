# src/constructor

Директория `src/constructor` содержит код для графического интерфейса (GUI) приложения. Использует `PySide6` и `qfluentwidgets` для создания окон, виджетов и взаимодействия с пользователем.

## Структура

* `gui_app.py`: Точка входа для GUI. Создаёт `QApplication` и запускает `GUIController`.
* `gui_controller.py`: Центральный контроллер GUI, координирует работу виджетов.
* `main_window.py`: Основное окно приложения, построенное на `FluentWindow`.
* `widgets/`: Виджеты интерфейса (например, `SheetEditor`, `ProjectExplorer`).
* `__init__.py`: Инициализация пакета `constructor`.

## Основные компоненты

* **GUIController**: Управляет жизненным циклом GUI.
* **MainWindow**: Основное окно.
* **SheetEditor**: Редактор данных листа.
* **ProjectExplorer**: Обозреватель проекта и листов.
