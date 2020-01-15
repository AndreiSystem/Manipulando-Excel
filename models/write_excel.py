import xlsxwriter
import os


# ------------------------------
#  Manipulando Excel via script
#  Autor : Andrei Gustavo
#  Data : (14/01/2020)
# ------------------------------


def check_if_the_path_of_the_file() -> str:
    route = 'C:/Users/andre/PycharmProjects/Excel/archives_excel/'
    file = route + 'excel.xlsx'

    if not os.path.exists(route):
        os.makedirs(route)

    if not os.path.exists(file):
        raise Exception('NonexistentFile!')

    return file


def excel_writing_function(list_of_dict: list):
    path = check_if_the_path_of_the_file()
    outWorkbook = xlsxwriter.Workbook(path)
    outSheet = outWorkbook.add_worksheet()

    # write headers
    col = 0
    headers = _receive_header_keys(list_of_dict)
    for key in headers:
        outSheet.write(0, col, key)
        col += 1

    # write data to file
    data_of_users = _receive_data_of_users(list_of_dict)
    index = 0
    for row in range(1, len(list_of_dict) + 1):
        for col in range(len(headers)):
            outSheet.write(row, col, data_of_users[index])
            index += 1


    outWorkbook.close()



def _receive_data_of_users(list: list) -> list:
    data_users = []
    for dict in list:
        for values in dict.values():
            data_users.append(values)

    return data_users

def _receive_header_keys(list: list) -> list:
    existing_keys = []
    for dict in list:
        for key in dict.keys():
            if key not in existing_keys:
                existing_keys.append(key)

    return existing_keys


list_users_data = [{'Name': 'Lucas', 'LastName': 'Silva', 'Age': 17, 'Country': 'Gaspar'}, {'Name': 'Douglas', 'LastName': 'Ronchi', 'Age': 27, 'Country': 'Blumenau'}, {'Name': 'Andrei', 'LastName': 'Teixeira da Luz', 'Age': 19, 'Country': 'Ilhota'}]
excel_writing_function(list_users_data)