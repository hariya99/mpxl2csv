import unittest
import sys, os 
sys.path.append(os.path.abspath(os.path.join('..', 'mpxl2csv')))
from mpxl2csv import Excel2Csv


class TestMP(unittest.TestCase):

    def setUp(self) -> None:
        self.mp = Excel2Csv()
    
    def tearDown(self) -> None:
        del self.mp 

    def test_warning(self):
        with self.assertWarns(UserWarning):
            mp = Excel2Csv(num_processes=9)
        
    def test_list_xl_files(self):
        path = "../sample_data/"
        xl_files = self.mp._get_xl_files(path)
        self.assertEqual(len(xl_files), 3)


if __name__ == '__main__':
    unittest.main()