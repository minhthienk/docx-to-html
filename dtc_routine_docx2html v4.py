from sys import exit

from dtc_routine_docx2html_functions import routine_final_convert
from dtc_routine_docx2html_functions import capitalize_each_word


import re

path = r'C:\Users\ThienNguyen\Desktop\P0135.docx'
routine_final_convert(path=path, loop=True)

