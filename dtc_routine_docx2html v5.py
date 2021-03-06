from sys import exit

from dtc_routine_docx2html_functions import routine_final_convert
from dtc_routine_docx2html_functions import capitalize_each_word


import re


path=r'F:\Dropbox\DTC Routine Project\Thien\Finished\1997 Honda Accord L4, 2.2L\Diagnostic Routine\DTC\P0141\P0141.docx'
routine_final_convert(path=path, input_interface_flag=True)

