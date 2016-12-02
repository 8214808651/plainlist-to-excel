# plainlist-to-excel
Formats a plaintext list into a more printer-friendly numbered excel file

# Usage
Install xlwings:
  pip install xlwings

Run listToExcel:
  python listToExcel inputFile outputFile
  
Note: The input file must be a list with one word per line. The output file must be a valid excel file.

#TODO

* Stop it from breaking at >13 columns per page
* Create an output file if one doesn't already exist
* Add error handling for invalid input files
