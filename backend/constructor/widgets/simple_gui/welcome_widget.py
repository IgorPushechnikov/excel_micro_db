# backend/constructor/widgets/simple_gui/welcome_widget.py
"""
Приветственный экран для упрощённого GUI.
"""
from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton
from PySide6.QtCore import Qt, Signal


class WelcomeWidget(QWidget):
    """Приветственный экран с кнопкой импорта."""
    
    import_requested = Signal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        
        label = QLabel("Добро пожаловать в Excel Micro DB!\n\n"
                      "Нажмите кнопку ниже для импорта и анализа Excel-файла.")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font-size: 16px; padding: 20px;")
        layout.addWidget(label)
        
        import_btn = QPushButton("Импорт Excel-файла")
        import_btn.setStyleSheet("font-size: 14px; padding: 10px;")
        import_btn.clicked.connect(self.import_requested.emit)
        layout.addWidget(import_btn)
        
        self.setLayout(layout)