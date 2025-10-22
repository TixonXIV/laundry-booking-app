from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import psycopg2
from psycopg2.extras import DictCursor
from datetime import datetime
from docx import Document
from io import BytesIO
import os

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'super_secret_key')

# Глобальное определение time_sort_key
def time_sort_key(ts):
    try:
        start_time = ts.split('-')[0]
        h, m = map(int, start_time.split(':'))
        return h * 60 + m
    except:
        return 0

# Получение соединения с PostgreSQL
def get_db_connection():
    return psycopg2.connect(os.getenv('DATABASE_URL'), cursor_factory=DictCursor)

# Инициализация базы данных
def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (id SERIAL PRIMARY KEY,
                  surname TEXT NOT NULL,
                  room TEXT NOT NULL,
                  UNIQUE(surname, room))''')
    c.execute('''CREATE TABLE IF NOT EXISTS days
                 (name TEXT PRIMARY KEY,
                  order_num INTEGER UNIQUE)''')
    c.execute('''CREATE TABLE IF NOT EXISTS machines
                 (number SERIAL PRIMARY KEY,
                  status TEXT DEFAULT 'active')''')
    c.execute('''CREATE TABLE IF NOT EXISTS slots
                 (id SERIAL PRIMARY KEY,
                  day_name TEXT NOT NULL,
                  time_slot TEXT NOT NULL,
                  machine_number INTEGER NOT NULL,
                  user_id INTEGER,
                  FOREIGN KEY(day_name) REFERENCES days(name),
                  FOREIGN KEY(machine_number) REFERENCES machines(number),
                  UNIQUE(day_name, time_slot, machine_number))''')
    conn.commit()
    c.close()
    conn.close()

init_db()

# Стандартные временные слоты
default_time_slots = [
    '7:00-9:00',
    '10:00-12:00',
    '13:00-15:00',
    '16:00-18:00',
    '19:00-21:00',
    '22:00-24:00'
]

def get_days():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('SELECT name FROM days ORDER BY order_num')
    days = [row['name'] for row in c.fetchall()]
    c.close()
    conn.close()
    return days

def get_machines():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('SELECT number, status FROM machines ORDER BY number')
    machines = [(row['number'], row['status']) for row in c.fetchall()]
    c.close()
    conn.close()
    return machines

def get_slots(for_admin=False):
    conn = get_db_connection()
    c = conn.cursor()
    if for_admin:
        c.execute('SELECT s.day_name, s.time_slot, s.machine_number, s.user_id, m.status FROM slots s LEFT JOIN machines m ON s.machine_number = m.number ORDER BY s.day_name, time_slot, s.machine_number')
    else:
        c.execute('SELECT s.day_name, s.time_slot, s.machine_number, s.user_id, m.status FROM slots s LEFT JOIN machines m ON s.machine_number = m.number WHERE m.status = "active" ORDER BY s.day_name, time_slot, s.machine_number')
    slots = c.fetchall()
    c.close()
    conn.close()
    return slots

def is_user_booked(user_id):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('SELECT * FROM slots WHERE user_id = %s', (user_id,))
    booked = c.fetchone()
    c.close()
    conn.close()
    return booked is not None

@app.route('/', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        surname = request.form['surname']
        room = request.form['room']
        conn = get_db_connection()
        c = conn.cursor()
        try:
            c.execute('INSERT INTO users (surname, room) VALUES (%s, %s) RETURNING id', (surname, room))
            user_id = c.fetchone()['id']
            conn.commit()
            session['user_id'] = user_id
            session['surname'] = surname
            session['room'] = room
            c.close()
            conn.close()
            return redirect(url_for('schedule'))
        except psycopg2.IntegrityError:
            c.execute('SELECT id FROM users WHERE surname = %s AND room = %s', (surname, room))
            user = c.fetchone()
            if user:
                session['user_id'] = user['id']
                session['surname'] = surname
                session['room'] = room
                c.close()
                conn.close()
                return redirect(url_for('schedule'))
            else:
                flash('Ошибка регистрации')
        c.close()
        conn.close()
    return render_template('register.html')

@app.route('/schedule', methods=['GET', 'POST'])
def schedule():
    if 'user_id' not in session:
        return redirect(url_for('register'))
    
    if is_user_booked(session['user_id']):
        flash('Вы уже записаны. Только одна запись на пользователя.')
    
    days = get_days()
    machines = get_machines()
    active_machines = [m for m, s in machines if s == 'active']
    num_machines = len(active_machines)
    
    slots = get_slots(for_admin=False)
    current_time_slots = set(ts for _, ts, _, _, _ in slots)
    time_slots = sorted(list(current_time_slots), key=time_sort_key)
    
    schedule_data = {day: {ts: [None] * num_machines for ts in time_slots} for day in days}
    machine_map = {active_machines[i]: i for i in range(num_machines)}
    
    for day, ts, machine, user_id, status in slots:
        if day in schedule_data and ts in schedule_data[day]:
            idx = machine_map.get(machine)
            if idx is not None:
                if user_id:
                    conn = get_db_connection()
                    c = conn.cursor()
                    c.execute('SELECT surname, room FROM users WHERE id = %s', (user_id,))
                    user = c.fetchone()
                    c.close()
                    conn.close()
                    schedule_data[day][ts][idx] = f"{user['surname']} {user['room']}"
                else:
                    schedule_data[day][ts][idx] = None
    
    if request.method == 'POST':
        if is_user_booked(session['user_id']):
            return redirect(url_for('schedule'))
        
        day = request.form['day']
        ts = request.form['time_slot']
        machine = int(request.form['machine'])
        
        conn = get_db_connection()
        c = conn.cursor()
        c.execute('''SELECT s.user_id, m.status FROM slots s 
                     JOIN machines m ON s.machine_number = m.number 
                     WHERE s.day_name = %s AND s.time_slot = %s AND s.machine_number = %s''', (day, ts, machine))
        slot = c.fetchone()
        if slot and slot['user_id'] is None and slot['status'] == 'active':
            c.execute('UPDATE slots SET user_id = %s WHERE day_name = %s AND time_slot = %s AND machine_number = %s', (session['user_id'], day, ts, machine))
            conn.commit()
            flash('Запись успешна!')
        else:
            flash('Слот занят, не существует или машина отключена.')
        c.close()
        conn.close()
        return redirect(url_for('schedule'))
    
    return render_template('schedule.html', schedule_data=schedule_data, days=days, time_slots=time_slots, machines=active_machines)

@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():
    if session.get('admin'):
        return redirect(url_for('admin'))
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username == 'admin' and password == 'admin123':
            session['admin'] = True
            return redirect(url_for('admin'))
        else:
            flash('Неверный логин или пароль')
    return render_template('admin_login.html')

@app.route('/admin_logout')
def admin_logout():
    session.pop('admin', None)
    return redirect(url_for('admin_login'))

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if not session.get('admin'):
        return redirect(url_for('admin_login'))
    
    days = get_days()
    machines = get_machines()
    num_machines = len(machines)
    
    slots = get_slots(for_admin=True)
    current_time_slots = set(ts for _, ts, _, _, _ in slots)
    time_slots = sorted(list(current_time_slots), key=time_sort_key)
    
    schedule_data = {day: {ts: [None] * num_machines for ts in time_slots} for day in days}
    machine_map = {machines[i][0]: i for i in range(num_machines)}
    
    for day, ts, machine, user_id, status in slots:
        if day in schedule_data and ts in schedule_data[day]:
            idx = machine_map.get(machine)
            if idx is not None:
                if status == 'disabled':
                    schedule_data[day][ts][idx] = 'Отключена'
                elif user_id:
                    conn = get_db_connection()
                    c = conn.cursor()
                    c.execute('SELECT surname, room FROM users WHERE id = %s', (user_id,))
                    user = c.fetchone()
                    c.close()
                    conn.close()
                    schedule_data[day][ts][idx] = f"{user['surname']} {user['room']}"
                else:
                    schedule_data[day][ts][idx] = None
    
    if request.method == 'POST':
        action = request.form.get('action')
        conn = get_db_connection()
        c = conn.cursor()
        
        if action == 'reset':
            c.execute('UPDATE slots SET user_id = NULL')
            conn.commit()
            flash('Таблица сброшена.')
        
        elif action == 'factory_reset':
            c.execute('DELETE FROM slots')
            conn.commit()
            days = get_days()
            machines = [m for m, _ in get_machines()]
            for day in days:
                for ts in default_time_slots:
                    for machine in machines:
                        c.execute('INSERT INTO slots (day_name, time_slot, machine_number, user_id) VALUES (%s, %s, %s, NULL)', (day, ts, machine))
            conn.commit()
            flash('Таблица возвращена к заводским настройкам.')
        
        # ... (остальные actions как в вашем коде, с %s вместо ? и RETURNING где нужно)
        # Для примера, add_time:
        elif action == 'add_time':
            new_time = request.form['new_time']
            if new_time not in current_time_slots:
                for day in days:
                    for machine, _ in machines:
                        c.execute('INSERT INTO slots (day_name, time_slot, machine_number, user_id) VALUES (%s, %s, %s, NULL)', (day, new_time, machine))
                conn.commit()
                flash('Новый временной слот добавлен.')
        
        # Для edit:
        elif action == 'edit':
            day = request.form['day']
            ts = request.form['time_slot']
            machine = int(request.form['machine'])
            new_value = request.form['new_value'].strip()
            
            c.execute('SELECT status FROM machines WHERE number = %s', (machine,))
            status = c.fetchone()
            if status and status['status'] == 'disabled':
                flash('Машина отключена, нельзя редактировать.')
            else:
                if new_value == '':
                    c.execute('UPDATE slots SET user_id = NULL WHERE day_name = %s AND time_slot = %s AND machine_number = %s', (day, ts, machine))
                    conn.commit()
                    flash('Запись пользователя удалена.')
                else:
                    try:
                        surname, room = new_value.split(' ')
                        c.execute('SELECT id FROM users WHERE surname = %s AND room = %s', (surname, room))
                        user = c.fetchone()
                        if not user:
                            c.execute('INSERT INTO users (surname, room) VALUES (%s, %s) RETURNING id', (surname, room))
                            user_id = c.fetchone()['id']
                        else:
                            user_id = user['id']
                        c.execute('UPDATE slots SET user_id = %s WHERE day_name = %s AND time_slot = %s AND machine_number = %s', (user_id, day, ts, machine))
                        conn.commit()
                        flash('Слот отредактирован.')
                    except ValueError:
                        flash('Неверный формат: Фамилия Комната')
        
        c.close()
        conn.close()
        return redirect(url_for('admin'))
    
    return render_template('admin.html', schedule_data=schedule_data, days=days, time_slots=time_slots, machines=[m for m, _ in machines])

# ... (остальные маршруты как в вашем коде, с аналогичными заменами для psycopg2)

if __name__ == '__main__':
    app.run(debug=True)