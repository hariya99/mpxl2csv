import multiprocessing
from concurrent.futures import ProcessPoolExecutor
import warnings
from pathlib import Path
from typing import Generator 
import openpyxl as opx
import xlrd
# import time

class Excel2Csv(object):
    """
        Usage:
            from mpxl2csv import Excel2Csv

            mp = Excel2Csv(num_processes=3, delimiter="|")
            xl_path = os.path.join(os.getcwd(), "sample_input")
            csv_path = os.path.join(os.getcwd(), "sample_output")
            mp.convert(xl_path, csv_path) 
    """

    def __init__(self, num_processes : int = 2, 
                 delimiter : str = "|") -> None:
        """
            Utilize Python's multiprocessing module to convert .xlsx, xls files to .csv files.
            Param : num_processes: Number of processes to use for conversion
            Param : delimiter: Delimiter to use for .csv files
        """
        available_cores : str  = multiprocessing.cpu_count()
        if num_processes > available_cores:
            warnings.warn(f"Number of processes ({num_processes}) " + 
                        f"is greater than the number of CPUs ({available_cores}). " +
                        f"It is advisable to use a number of processes less " +
                        f"than or equal to the number of CPUs.")
        self.num_processes : int = num_processes
        self.delimiter : str = delimiter
        self._extns : tuple = (".xlsx", ".xls")


    def convert(self, xl_base_path : str, csv_base_path : str) -> None:
        """
            Convert all .xlsx files in a directory to .csv files.
        """
        xl_files : Generator[Path, None, None] = self._get_xl_files(xl_base_path)
        self._convert_files(xl_files, csv_base_path)

    def _get_xl_files(self, xl_base_path : str) -> Generator[Path, None, None]:
        """
            Get all .xlsx and .xls files from a directory.
            The directory can have subdirectories.
            return a generator object
        """
        return (path_obj for path_obj in 
                Path(xl_base_path).rglob("*")
                if path_obj.suffix in self._extns)

    def _get_csv_files(self, csv_base_path : str) -> Generator[Path, None, None]:
        """
            Get all .csv files from a directory.
            The directory can have subdirectories.
            return a generator object
        """
        return (path_obj for path_obj in Path(csv_base_path).rglob("*.csv"))
    
    def _convert_files(self, xl_files : Generator[Path, None, None], 
                       csv_base_path : str) -> None:
        """
            Convert all .xlsx and .xls files to .csv files.
        """
        # Method - 1 : Using multiprocessing module
        # processes = []
        # with multiprocessing.Pool(self.num_processes) as pool:
        #     for xl_file in xl_files:
        #         csv_file = Path(csv_base_path).joinpath(xl_file.stem + ".csv")
        #         processes.append(pool.apply_async(self._convert_file, 
        #                                           args=(str(xl_file), str(csv_file))))
        #     for process in processes:
        #         process.get()

        # Method - 2 : Using concurrent.futures module
        with ProcessPoolExecutor(max_workers=self.num_processes) as executor:
            for xl_file in xl_files:
                csv_file = Path(csv_base_path).joinpath(xl_file.stem + ".csv")
                if xl_file.suffix == self._extns[0]:
                    executor.submit(self._convert_file_opx, str(xl_file), str(csv_file))
                else: 
                    executor.submit(self._convert_file_xlrd, str(xl_file), str(csv_file))

    
    def _convert_file_opx(self, xl_path : str, csv_path : str )-> None:
        """
            Convert a single .xlsx file to a .csv file using openpyxl 
        """
        # print(f"Conversion started for file: {xl_path}")
        # start_time = time.time()
        wb = opx.load_workbook(xl_path, read_only=True, data_only=True)
        ws = wb.active
        with open(csv_path, "w", encoding="utf-8") as f:
            for row in ws.rows:
                row_values = (str(cell.value) for cell in row)
                f.write(self.delimiter.join(row_values) + "\n")
        
        wb.close()
        # end_time = time.time()
        # print(f"Time taken to convert to {csv_path} is {end_time - start_time} seconds")

    def _convert_file_xlrd(self,xl_path : str, csv_path : str )-> None:
        """
            Convert a single .xls file to a .csv file using xlrd 
        """
        # print(f"Conversion started for file: {xl_path}")
        # start_time = time.time()
        wb = xlrd.open_workbook(xl_path)
        ws = wb.sheet_by_index(0)
        with open(csv_path, "w", encoding="utf-8") as f:
            for row_num in range(ws.nrows):
                row_values = (str(cell.value) for cell in ws.row(row_num))
                f.write(self.delimiter.join(row_values) + "\n")
        wb.release_resources()

        # end_time = time.time()
        # print(f"Time taken to convert to {csv_path} is {end_time - start_time} seconds")
