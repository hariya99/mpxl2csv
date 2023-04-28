import unittest
from typing import Generator
from pathlib import Path
import os, sys 
sys.path.append(os.path.abspath(os.path.join('..', 'mpxl2csv')))
from mpxl2csv import Excel2Csv


class TestMP(unittest.TestCase):

    def setUp(self) -> None:
        self.mp = Excel2Csv()
        self.input_path : str = str(Path(".").joinpath("resources", "sample_input"))
        self.output_path : str = str(Path(".").joinpath("resources", "sample_output"))
    
    def tearDown(self) -> None:
        del self.mp 

    def test_warning(self):
        with self.assertWarns(UserWarning):
            mp = Excel2Csv(num_processes=9)
    
    def test_list_xl_files(self):
        xl_files = self.mp._get_xl_files(self.input_path)
        self.assertIsInstance(xl_files, Generator)

    def test_input_file_type(self):
        xl_files = self.mp._get_xl_files(self.input_path)
        for xl_file in xl_files:
            self.assertIsInstance(xl_file, Path)
            self.assertRegex(xl_file.suffix, r"\.xlsx|\.xls")
    
    def test_ouput_file_type(self):
        self.mp.convert(self.input_path, self.output_path)
        csv_files = self.mp._get_csv_files(self.output_path)
        for csv_file in csv_files:
            self.assertIsInstance(csv_file, Path)
            self.assertEqual(csv_file.suffix, ".csv")


if __name__ == '__main__':
    unittest.main()