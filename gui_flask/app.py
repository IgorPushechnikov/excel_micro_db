# gui_flask/app.py

from flask import Flask, render_template

# Создание экземпляра Flask-приложения
# Указываем, что шаблоны находятся в папке ../gui_flask/templates относительно этого файла
app = Flask(__name__, template_folder='../gui_flask/templates', static_folder='../gui_flask/static')

@app.route('/')
def index():
    """
    Маршрут для главной страницы GUI.
    """
    # Отображаем базовый шаблон index.html
    return render_template('index.html')

# Проверка, что скрипт запущен напрямую, а не импортирован
if __name__ == '__main__':
    # Запуск Flask-приложения в режиме разработки
    # host='0.0.0.0' позволяет получить доступ с других устройств в сети (опционально)
    # port=5000 стандартный для Flask
    app.run(debug=True, host='0.0.0.0', port=5000)
