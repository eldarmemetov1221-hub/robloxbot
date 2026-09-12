import sqlite3

# Подключение к базе
conn = sqlite3.connect("codes.db")
cursor = conn.cursor()

# Посмотреть все коды
cursor.execute("SELECT code, product, status, instruction FROM codes")
rows = cursor.fetchall()

print("=== Список кодов в базе ===")
for row in rows:
    print(f"Код: {row[0]} | Товар: {row[1]} | Статус: {row[2]} | Инструкция: {row[3]}")

conn.close()
