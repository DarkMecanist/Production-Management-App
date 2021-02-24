import sqlite3

# def connect(func):
#     def wrap(*args):
#         conn = sqlite3.connect('Database.db', isolation_level=None)
#         cursor = conn.cursor()
#         func(*args, cursor)
#         conn.close()
#
#     return wrap


def load_order_parts_from_database(mode='all', material_ref=None, rowid=None, *args):
    if mode == 'all':
        cursor.execute('''select *, rowid from order_parts order by client, order_num_client, name''')
        return cursor.fetchall()
    elif mode == 'new_task':
        cursor.execute('''select name, rowid from order_parts where material_ref = :material_ref order by client''', {'material_ref': material_ref})
        return cursor.fetchall()
    elif mode == 'rowid':
        cursor.execute('''select * from order_parts where rowid = :rowid''', {'rowid': rowid})
        return cursor.fetchall()
    elif mode == 'produced_quantity':
        cursor.execute('select produced_quantity from order_parts where rowid = :rowid', {'rowid': rowid})
        return cursor.fetchone()


def remove_order_part_database(rowid):
    with conn:
        cursor.execute('''delete from order_parts where rowid = :rowid''', {'rowid': rowid})


def load_tasks_from_database(rowid=None):
    if rowid == None:
        cursor.execute('select *, rowid from tasks order by priority')
    else:
        cursor.execute('''select * from tasks where rowid = :rowid''', {'rowid': rowid})
    return cursor.fetchall()


def remove_task_database(rowid):
    with conn:
        cursor.execute('delete from tasks where rowid = :rowid', {'rowid': rowid})


def create_tasks_table(*args):
    cursor.execute('''create table if not exists tasks(
                        current_path text,
                        original_path text,
                        machine text,
                        material_ref text,
                        notes text,
                        estimated_sheets_required float,
                        estimated_time int,
                        start_date text,
                        end_date text,
                        priority int,
                        order_parts text,
                        aggregated_tasks text,
                        aggregated_index text
                        )''')


def delete_tasks_table(*args):
    cursor.execute('drop table if exists tasks')


def delete_order_parts_table(*args):
    cursor.execute('drop table if exists order_parts')


def create_order_parts_table(*args):
    cursor.execute('''create table if not exists order_parts(
                        ref text,
                        name text,
                        material_ref text,
                        quantity int,
                        produced_quantity int,
                        order_num text,
                        order_num_client text,
                        client text,
                        date_modified text,
                        due_date text,
                        additional_operations boolean
                        )''')


def create_parts_table(*args):
    cursor.execute('''create table if not exists parts(
                        ref text,
                        name text,
                        weight float,
                        material_ref text,
                        time int,
                        client text,
                        date_modified text
                        )''')


def delete_parts_table(*args):
    cursor.execute('drop table if exists parts')


def insert_new_part(ref, name, weight, material_ref, time, client, date_modified, *args):
    cursor.execute('insert into parts values(:ref, :name, :weight, :material_ref, :time, :client, :date_modified)',
                     {'ref': ref,
                      'name': name,
                      'weight': weight,
                      'material_ref': material_ref,
                      'time': time,
                      'client': client,
                      'date_modified': date_modified
                      })


def get_parts(*args):
    cursor.execute('''select *, rowid from parts order by name asc''')
    return cursor.fetchall()


def create_materials_table(*args):
    '''Creates a table(if not already created) which can hold task information between user sessions'''

    cursor.execute('''create table if not exists materials(
                    ref text,
                    type text,
                    spec text,
                    thickness float,
                    length float,
                    width float,
                    density float,
                    stock_weight float,
                    stock_num_sheets float,
                    min_stock int,
                    date_modified text,
                    client text)''')


def insert_new_material(ref, type, spec, thickness, length, width, density, stock_weight, stock_num_sheets, min_stock, date_modified, client, *args):
    cursor.execute('insert into materials values(:ref, :type, :spec, :thickness, :length, :width, :density, :stock_weight, :stock_num_sheets, :min_stock, :date_modified, :client)',
                     {'ref': ref,
                      'type': type,
                      'spec': spec,
                      'thickness': thickness,
                      'length': length,
                      'width': width,
                      'density': density,
                      'stock_weight': stock_weight,
                      'stock_num_sheets': stock_num_sheets,
                      'min_stock': min_stock,
                      'date_modified': date_modified,
                      'client': client
                      })


def get_material_ref(type, spec, thickness):
    cursor.execute('''select ref from materials where type = :type and spec = :spec and thickness = :thickness''', {'type': type, 'spec': spec, 'thickness': thickness})
    return cursor.fetchone()[0]


def get_materials(*args):
    cursor.execute('''select *, rowid from materials order by thickness asc''')
    return cursor.fetchall()


def create_shifts_table(*args):
    cursor.execute('''create table if not exists shifts(
                    machine text,
                    time_start text,
                    time_finish text,
                    time_break text,
                    break_duration int
                      )''')


def insert_new_shift(machine, time_start, time_finish, time_break, break_duration):
    cursor.execute('''insert into shifts values(:machine, :time_start, :time_finish, :time_break, :break_duration)''',
                   {'machine': machine,
                    'time_start': time_start,
                    'time_finish': time_finish,
                    'time_break': time_break,
                    'break_duration': break_duration})


def load_shifts(machine=None):
    if machine == None:
        query = '''select * from shifts'''
    else:
        query = '''select * from shifts where machine = :machine''', {'machine': machine}

    cursor.execute(query)

    return cursor.fetchall()


def change_value_shifts_database(value, field, rowid):
    with conn:
        if field == 'time_start':
            cursor.execute('update shifts set time_start = :time_start where rowid = :rowid', {'time_start': value, 'rowid': rowid})
        elif field == 'time_finish':
            cursor.execute('update shifts set time_finish = :time_finish where rowid = :rowid', {'time_finish': value, 'rowid': rowid})
        elif field == 'time_break':
            cursor.execute('update shifts set time_break = :time_break where rowid = :rowid', {'time_break': value, 'rowid': rowid})
        elif field == 'break_duration':
            cursor.execute('update shifts set break_duration = :break_duration where rowid = :rowid', {'break_duration': value, 'rowid': rowid})


def remove_shifts(rowid):
    with conn:
        cursor.execute('delete from tasks where rowid = :rowid', {'rowid': rowid})




if __name__ == '__main__':
    conn = sqlite3.connect('Database.db', isolation_level=None)
    cursor = conn.cursor()

    create_shifts_table()

    # insert_new_shift('LC5', '6:00', '14:00', '10:00', 30)
    # insert_new_shift('LC5', '14:00', '22:00', '18:00', 30)
    # insert_new_shift('LF3015', '6:00', '14:00', '10:00', 30)
    # insert_new_shift('LF3015', '14:00', '22:00', '18:00', 30)

    # shifts = load_shifts()
    #
    # for shift in shifts:
    #     print(shift)

    conn.close()