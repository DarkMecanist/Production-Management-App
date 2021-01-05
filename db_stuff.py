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


if __name__ == '__main__':
    conn = sqlite3.connect('Database.db', isolation_level=None)
    cursor = conn.cursor()

    # # remove_task_database(18)
    #
    # task_list = load_tasks_from_database()
    # for task in task_list:
    #     print(task)

    order_list = load_order_parts_from_database()
    for order in order_list:
        if order[7] == 'RST - CONSTRUTORA DE MAQUINAS E ACESSORIOS, SA':
            remove_order_part_database(order[11])
        print(order)

    conn.close()