import sqlite3
import datetime
import openpyxl

current_date = datetime.date.today()
class Application:
    def __init__(self, database):
        self.conn = sqlite3.connect(database)
        self.c = self.conn.cursor()
        self.c.execute('''CREATE TABLE IF NOT EXISTS events 
                          (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                           name TEXT, 
                           date TEXT, 
                           description TEXT,
                           notification_sent BOOLEAN DEFAULT 0)''')
        self.conn.commit()

    def add_event(self, name, date, description):
        # Проверка на прошедшую дату
        if date < datetime.date.today():
            print("Ошибка: нельзя добавить мероприятие с прошедшей датой")
            return

        # Добавление информации о мероприятии
        self.c.execute("INSERT INTO events (name, date, description) VALUES (?, ?, ?)", (name, date, description))
        self.conn.commit()

    def get_upcoming_events(self):
        # Получение списка предстоящих мероприятий
        today = datetime.date.today()
        self.c.execute("SELECT * FROM events WHERE date = ? AND notification_sent = 0", (today,))
        upcoming_events = self.c.fetchall()

        if len(upcoming_events) == 0:
            print("Запланированных событий нет")

        return []

    def send_notification(self):
        events = self.get_upcoming_events()
        for event in events:
            message = f"Напоминаем, что мероприятие \"{event[1]}\" начнется {event[2]}. {event[3]}"
            print(message)
            # Пометка, что уведомление отправлено
            self.c.execute("UPDATE events SET notification_sent = 1 WHERE id = ?", (event[0],))
            self.conn.commit()

    def export_past_events(self):
        # Выгрузка всех прошедших мероприятий в файл xlsx
        past_events = []
        today = datetime.date.today()
        self.c.execute("SELECT id, name, date, description FROM events WHERE date < ?", (today,))
        past_events = self.c.fetchall()

        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(["ID", "Название", "Дата", "Описание"])

        for event in past_events:
            worksheet.append(event)

        workbook.save('past_events.xlsx')

    def close_connection(self):
        # Закрытие соединения с базой данных
        self.conn.close()


# Пример использования

# Создание экземпляра приложения
app = Application('events.db')

# Добавление мероприятий
app.add_event("Встреча с клиентом", datetime.date(2024, 1, 10), "Встреча в офисе")
app.add_event("Презентация", datetime.date(2024, 2, 10), "Презентация нового продукта")

# Получение списка предстоящих мероприятий
upcoming_events = app.get_upcoming_events()
for event in upcoming_events:
    print(f"{event[1]} - {event[2]}")

# Отправка уведомления о старте мероприятия
app.send_notification()

# Выгрузка всех прошедших мероприятий в файл CSV
app.export_past_events()

# Закрытие соединения с базой данных
app.close_connection()