import psycopg2
from psycopg2 import Error
import openpyxl
import argparse

parser = argparse.ArgumentParser(description='Input path to the file')
parser.add_argument('--xls_path', help='Full path to the file', type=str)

args = parser.parse_args()
if args.xls_path:
    print(args.xls_path)
else:
    print("Please, input correct value of path_argument")

def main():
    try:
        # Подключаемся к существующей базе данных
        connection = psycopg2.connect(user="postgres",
                                      password="???????",
                                      host="127.0.0.1",
                                      port="5432",
                                      database="for_tests")

        cursor = connection.cursor()

        upload = "INSERT INTO endpoint_names (endpoint_id, endpoint_name) " \
                 "VALUES (%s,%s)"
        update_name = "UPDATE endpoint_names SET endpoint_name=%s WHERE " \
                    "endpoint_id=%s"

        book = openpyxl.open(args.xls_path, read_only=True)
        sheet = book.active
        names = {}

        for row in range(2, sheet.max_row+1):
            if sheet[row][0].value:
                names[sheet[row][0].value] = sheet[row][1].value

        for p_id, name in names.items():
            cursor.execute("SELECT endpoint_id FROM endpoint_names WHERE "
                           "endpoint_id=%s", (p_id,))
            if cursor.fetchone():
                cursor.execute(update_name, (name, p_id))
                connection.commit()
            else:
                cursor.execute(upload, (p_id, name))
                connection.commit()

    except (Exception, Error) as err:
        print("Ошибка при работе с PostgreSQL", err)
    finally:
        if connection:
            cursor.close()
            connection.close()
            print("Соединение с PostgreSQL закрыто")



if __name__ == '__main__':
    main()
    input("Нажмите Enter, чтобы закрыть окно...")