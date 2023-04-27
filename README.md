# A Python Multiprocessing library to convert Excel(xlsx) files to csv
### Python based libraries are notoriously slow to process large Excel files. This library utilizes Python multiprocessing and openpyxl to process multiple Excel files in parallel so as to reduce the total time taken to convert them. 

### Usage:
```
from mpxl2csv import Excel2Csv

mp = Excel2Csv(num_processes=3, delimiter="|")
xl_path = os.path.join(os.getcwd(), "sample_input")
csv_path = os.path.join(os.getcwd(), "sample_output")
mp.convert(xl_path, csv_path) 

```