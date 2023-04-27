import os, multiprocessing, warnings
from typing import List 
import openpyxl as opx


class Excel2Csv(object):
    """
        Usage:
            from mpxl2csv import Excel2Csv

            mp = Excel2Csv(num_processes=3, delimiter="|")
            xl_path = os.path.join(os.getcwd(), "sample_input")
            csv_path = os.path.join(os.getcwd(), "sample_output")
            mp.convert(xl_path, csv_path) 
    """

    def __init__(self, num_processes=2, 
                 engine="openpyxl",
                 delimiter = "|") -> None:
        """
            Utilize Python's multiprocessing module to convert .xlsx files to .csv files.
            Param : num_processes: Number of processes to use for conversion
            Param : engine: Engine to use for conversion. Currently only openpyxl is supported 
            Param : delimiter: Delimiter to use for .csv files
        """
        available_cores = multiprocessing.cpu_count()
        if num_processes > available_cores:
            warnings.warn(f"Number of processes ({num_processes}) " + 
                        f"is greater than the number of CPUs ({available_cores}). " +
                        f"It is advisable to use a number of processes less " +
                        f"than or equal to the number of CPUs.")
        self.num_processes = num_processes
        self.engine = engine
        self.delimiter = delimiter


    def convert(self, xl_base_path : str, csv_base_path : str) -> None:
        """
            Convert all .xlsx files in a directory to .csv files.
        """
        xl_files = self._get_xl_files(xl_base_path)
        self._convert_files(xl_files, csv_base_path)

    def _get_xl_files(self, xl_base_path : str) -> List[str]:
        """
            Get all .xlsx files in a directory.
            The directory can have subdirectories.
        """
        xl_files = []
        for root, dirs, files in os.walk(xl_base_path):
            for file in files:
                if file.endswith(".xlsx"):
                    xl_files.append(os.path.join(root, file))
        return xl_files 

    def _convert_files(self, xl_files : List[str], csv_base_path : str) -> None:
        """
            Convert all .xlsx files to .csv files.
        """
        processes = []
        with multiprocessing.Pool(self.num_processes) as pool:
            for xl_file in xl_files:
                csv_file = os.path.join(csv_base_path, os.path.basename(xl_file).replace(".xlsx", ".csv"))
                processes.append(pool.apply_async(self._convert_file, args=(xl_file, csv_file)))
            for process in processes:
                process.get()

    
    def _convert_file(self, xl_path : str, csv_path : str )-> None:
        """
            Convert a single .xlsx file to a .csv file using openpyxl 
        """
        wb = opx.load_workbook(xl_path, read_only=True, data_only=True)
        ws = wb.active
        with open(csv_path, "w") as f:
            for row in ws.rows:
                row_values = [str(cell.value) for cell in row]
                f.write(self.delimiter.join(row_values) + "\n")
        
        wb.close()

