import os
from app import app

# Инициализируем БД при первом запуске (нужно для Render/Railway — ephemeral filesystem)
db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'csupd.db')
if not os.path.exists(db_path):
    from init_db import main as init_db
    init_db()

if __name__ == '__main__':
    app.run()
