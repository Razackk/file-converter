from generic_converter import GenericConverter
from os.path import dirname, realpath, join
base_file_path = dirname(realpath("resources"))

if __name__ == '__main__':

    cv = GenericConverter()
    file1 = cv.excel_to_json(base_file_path + "/tests/resources/multiple-sheets/result.xlsx")
    file2 = cv.json_to_excel(base_file_path + "/tests/resources/json/fruit.json")
    file3 = cv.excel_to_csv(base_file_path + "/tests/resources/single-sheet/test.xlsx")
    dict_name = {"id": [10001], "name": ["Johnson"], "age": [25]}
    file4 = cv.dict_to_excel(dict_name)

    print(file1)
    # print(file2)
    # print(file3)
    # print(file4)
