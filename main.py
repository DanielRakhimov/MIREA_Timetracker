import sqlite3
import datetime as dt
import openpyxl

class Application:
    def __init__(self, database):
        self.conn = sqlite3.connect(database)
        self.c = self.conn.cursor()
        self.c.execute('''CREATE TABLE IF NOT EXISTS events 
                          (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                           name TEXT, 
                           datetime DATETIME, 
                           description TEXT,
                           notification_sent BOOLEAN DEFAULT 0)''')
        self.conn.commit()

    def add_event(self, name, datetime, description):
        # Проверка на прошедшую дату и время
        if datetime < dt.datetime.now():
            print("Ошибка: нельзя добавить мероприятие с прошедшей датой и временем")
            return

        # Добавление информации о мероприятии
        self.c.execute("INSERT INTO events (name, datetime, description) VALUES (?, ?, ?)", (name, datetime, description))
        self.conn.commit()

    def get_upcoming_events(self):
        # Получение списка предстоящих мероприятий в текущем дне
        now = dt.datetime.now()
        start_of_day = dt.datetime.combine(now.date(), dt.time())
        self.c.execute("SELECT * FROM events WHERE datetime >= ? AND datetime < ? AND notification_sent = 0",
                       (now, start_of_day + dt.timedelta(days=1)))
        upcoming_events = self.c.fetchall()

        if len(upcoming_events) == 0:
            print("Запланированных событий на сегодня нет")

        return upcoming_events

    def send_notification(self):
        events = self.get_upcoming_events()
        for event in events:
            event_datetime = str(event[2])[:-3]
            message = f"Напоминаем, что мероприятие \"{event[1]}\" начнется {event_datetime}. {event[3]}"
            print(message)
            # Пометка, что уведомление отправлено
            self.c.execute("UPDATE events SET notification_sent = 1 WHERE id = ?", (event[0],))
            self.conn.commit()

    def export_past_events(self):
        # Выгрузка всех прошедших мероприятий в файл xlsx
        past_events = []
        now = dt.datetime.now()
        self.c.execute("SELECT id, name, datetime, description FROM events WHERE datetime < ?", (now,))
        past_events = self.c.fetchall()

        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(["ID", "Название", "Дата и время", "Описание"])

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
app.add_event("Встреча с клиентом", dt.datetime(2024, 1, 12, 9, 0), "Встреча в офисе")
app.add_event("Презентация", dt.datetime(2024, 2, 12, 18, 40), "Презентация нового продукта")

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