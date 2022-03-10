import pytest
import os
from file_converter.converter.generic_converter import GenericConverter
from os.path import dirname, realpath, join
base_file_path = dirname(realpath(__file__)) + "/resources"


class TestConverters:
    cv = GenericConverter()

    @pytest.mark.one
    def test_excel_to_json(self):
        file_path = base_file_path + '/single-sheet/test.xlsx'
        sheet_name = 'Sheet1'
        result = '[{"id":10001,"name":"Johnson","age":25},{"id":10002,"name":"Razak","age":27}]'
        file = self.cv.excel_to_json(file_path, sheet_name)
        assert file == result

    @pytest.mark.two
    def test_json_to_excel(self):
        file_path = base_file_path + '/json/fruit.json'
        result = '           0'+os.linesep+'fruit  Apple'+os.linesep+'size   Large'+os.linesep+'color    Red'
        file = self.cv.json_to_excel(file_path)
        assert file == result

    @pytest.mark.three
    def test_excel_to_json_file_format(self):
        file_path = 'the.xlx'
        result = '{"Invalid": ["the.xlx file format or input is not supported"]}'
        file = self.cv.excel_to_json(file_path)
        assert file == result

    @pytest.mark.four
    def test_json_to_excel_file_format(self):
        file_path = 'the.jsn'
        result = 'the.jsn file format or input is not supported'
        file = self.cv.json_to_excel(file_path)
        assert file == result

    @pytest.mark.five
    def test_excel_handles_multiple_sheets(self):
        file_path = base_file_path + '/multiple-sheets/result.xlsx'
        result = '[{"sheet":"Sheet1","data":[{"Algorithm":"Random Forest","Accuracy":0.651263093,"F1":0.6382824211},{"Algorithm":"Gradient Boosting","Accuracy":0.6678989526,"F1":0.6535156608}]},{"sheet":"Sheet2","data":[{"Building Type":"BUNGALOW","A":0,"B":15},{"Building Type":"FLAT","A":2,"B":69}]},{"sheet":"Sheet3","data":[]}]'
        file = self.cv.excel_to_json(file_path)
        assert file == result

    @pytest.mark.six
    def test_excel_handles_multiple_sheets_2(self):
        file_path = base_file_path + '/single-sheet/test.xlsx'
        result = '[{"sheet":"Sheet1","data":[{"id":10001,"name":"Johnson","age":25},{"id":10002,"name":"Razak","age":27}]}]'
        file = self.cv.excel_to_json(file_path)
        assert file == result

    @pytest.mark.seven
    def test_excel_to_csv(self):
        file_path = base_file_path + '/single-sheet/test.xlsx'
        sheet_name = 'Sheet1'
        result = ',id,name,age'+os.linesep+'0,10001,Johnson,25'+os.linesep+'1,10002,Razak,27'+os.linesep
        file = self.cv.excel_to_csv(file_path, sheet_name)
        assert file == result

    @pytest.mark.eight
    def test_dict_to_excel(self):
        dict_name = {"id": [10001], "name": ["Johnson"], "age": [25]}
        result = '      id     name  age'+os.linesep+'0  10001  Johnson   25'
        file = self.cv.dict_to_excel(dict_name)
        assert file == result
