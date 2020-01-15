import os

def check_if_the_path_of_the_file() -> str:
    route = os.path.dirname(__file__) + '/archive_excel/'
    file = route + 'excel.xlsx'

    if not os.path.exists(file):
        raise Exception('fileNotFound')
    return file
