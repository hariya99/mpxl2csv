# MPXL2CSV
A Python Multiprocessing library to convert Excel(xlsx, xls) files to csv. Python based libraries are notoriously slow to process large Excel files. This library utilizes Python multiprocessing and openpyxl to process multiple Excel files in parallel so as to reduce the total time taken to convert them. 
# Installing
```
pip install mpxl2csv
```
# Sample Usage:
## Important Note: ```__main__``` guard is needed for Python's multiprocessing code to work. So it is advisable to wrap the code in a function as shown below and call it under the ```__main__``` guard.

```Python
import os 
from mpxl2csv import Excel2Csv

def main():
    mp = Excel2Csv(num_processes=3, delimiter="|")
    xl_base_path = os.path.join(os.getcwd(), "sample_input")
    csv_base_path = os.path.join(os.getcwd(), "sample_output")
    mp.convert(xl_base_path, csv_base_path) 

if __name__ == "__main__":
    main()
```