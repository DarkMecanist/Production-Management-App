import shutil, os, datetime

max_files = 5
direc = 'D:/PROJETOS/PROJETOS PYTHON/Projetos Trabalho/SOFTWARE PLANEAMENTO/Backups'


def create_db_copy():
    current_date = datetime.datetime.now()
    current_day = current_date.day
    current_month = current_date.month
    current_year = current_date.year

    shutil.copyfile('D:/PROJETOS/PROJETOS PYTHON/Projetos Trabalho/SOFTWARE PLANEAMENTO/Database.db', f'{direc}/{current_year}_{current_month}_{current_day}_Database.db')


def delete_old_backups():
    num_files = len(os.listdir(direc))

    if num_files == max_files:
        subdir = f'D:/PROJETOS/PROJETOS PYTHON/Projetos Trabalho/SOFTWARE PLANEAMENTO/Backups/{os.listdir(direc)[0]}'
        os.remove(subdir)

    create_db_copy()


if __name__ == '__main__':
    delete_old_backups()
