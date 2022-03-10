import pandas as pd
import json
from os.path import dirname, realpath


class GenericConverter:
    """
      A class used to convert excel to json and vice versa

      Methods
      -------
      excel_to_json()
          returns the json string or save json file
      json_to_excel()
          returns the json string or save json file
      """

    def excel_to_json(self, file_path: str, sheet_name: str = None, save_file: str = None):
        """
        Converts excel to JSon.

        If the argument `file_name` isn't passed is not an excel file or a string,
        invalid format will be returned.

        Parameters
        ----------
        file_path : str
        a string denoting the name of the file
        sheet_name : str
        the name of the sheet_name to be converted
        save_file: str = None
        an optional argument to save file, save_file = "y"

        Returns
        -------
        a json string
        """

        # check if filetype is a string and if file is valid
        if type(file_path) == str and file_path.endswith(".xlsx"):
            # check if single sheet or all sheets
            if sheet_name:
                # read the excel file
                excel_df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
            else:
                # read all sheets in excel file
                excel_dict = pd.read_excel(file_path, sheet_name=None, header=0)
                #  dictionary to list
                excel_dict_list = excel_dict.items()
                # dictionary list to dataframe
                excel_df = pd.DataFrame(excel_dict_list)
            # check if user wants file saved
            if save_file == "y":
                # remove excel extension from file name
                json_file_name = file_path.removesuffix(".xlsx")
                file_dir = dirname(realpath(file_path))
                # convert excel to json
                excel_df.to_json(file_dir + f"{json_file_name}.json")
                result = {"Success": [f"{file_path} converted to json and saved"]}
                return json.dumps(result)
            else:
                # user does not want file saved
                excel_df = excel_df.rename(columns={0: "sheet", 1: "data"})
                for i in range(len(excel_df)):
                    if excel_df.loc[i]["data"].empty:
                        excel_df.drop([i], axis=0, inplace=True)
                json_str = excel_df.to_json(orient="records")
                return json_str
        else:
            # user inputted an incorrect data type or file format
            invalid_format = {"Invalid": [f"{file_path} file format or input is not supported"]}
            return json.dumps(invalid_format)

    def excel_to_csv(self, file_path: str, sheet_name: str = None, save_file: str = None):
        """
        Converts excel to csv.

        If the argument `file_name` isn't passed is not an excel file or a string,
        invalid format will be returned.

        Parameters
        ----------
        file_path : str
        a string denoting the name of the file
        sheet_name : str
        the name of the sheet_name to be converted
        save_file: str = None
        an optional argument to save file, save_file = "y"

        Returns
        -------
        a csv string
        """

        # check if filetype is a string and if file is valid
        if type(file_path) == str and file_path.endswith(".xlsx"):
            # check if single sheet or all sheets
            if sheet_name:
                # read the excel file
                excel_df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
            else:
                # read all sheets in excel file
                excel_dict = pd.read_excel(file_path, sheet_name=None, header=0)
                #  dictionary to list
                excel_dict_list = excel_dict.items()
                # dictionary list to dataframe
                excel_df = pd.DataFrame(excel_dict_list)
            # check if user wants file saved
            if save_file == "y":
                # remove excel extension from file name
                json_file_name = file_path.removesuffix(".xlsx")
                # convert excel to csv
                excel_df.to_csv(f"{json_file_name}.csv")
                result = f"Success: {file_path} converted to csv and saved"
                return result
            else:
                # user does not want file saved
                csv_str = excel_df.to_csv(header=True)
                return csv_str
        else:
            # user inputted an incorrect data type or file format
            invalid_format = f"Invalid: {file_path} file format or input is not supported"
            return invalid_format

    def json_to_excel(self, file_path: str, save_file: str = None):

        """
        Converts JSon to excel.

        If the argument `file_name` isn't passed is not an excel file or a string,
        invalid format will be returned.

        Parameters
        ----------
        file_path : str
        a string denoting the name of the file
        save_file: str = None
        an optional argument to save file, save_file = "y"

        Returns
        -------
        an excel string
        """

        # check if filetype is a string and if file is valid
        if type(file_path) == str and file_path.endswith(".json"):
            # read the json file
            json_df = pd.read_json(file_path, orient='index')
            # check if user wants file saved
            if save_file == "y":
                # remove json extension from file name
                excel_file_name = file_path.removesuffix(".json")
                # convert json to excel
                json_df.to_excel(f"{excel_file_name}.xlsx")
                result = f"{file_path} converted to excel and saved"
                return result
            else:
                # user does not want file saved
                excel_str = str(json_df)
                return excel_str
        else:
            # user inputted an incorrect file
            result = f"{file_path} file format or input is not supported"
            return result

    def dict_to_excel(self, dict_name, save_file: str = None):
        """
        Converts JSon to excel.

        If the argument `dict_name` isn't passed is not an excel file or a string,
        invalid format will be returned.

        Parameters
        ----------
        dict_name : str
        a string denoting the name of the dictionary
        save_file: str = None
        an optional argument to save file, save_file = "y"

        Returns
        -------
        an excel string
        """

        # convert to dataframe
        dict_df = pd.DataFrame(data=dict_name)
        # check if user wants file saved
        if save_file == "y":
            # convert dictionary to excel
            file_dir = dirname(realpath(__file__))
            dict_df.to_excel(file_dir + f"{dict_name}.xlsx")
            result = f"{dict_name} converted to excel and saved"
            return result
        else:
            # user does not want file saved
            return str(dict_df)
