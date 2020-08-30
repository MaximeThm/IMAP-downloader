import os
import glob

path = '/Users/maximethomas/Desktop/Test'

file_list = []
for filename in glob.glob(os.path.join(path, '*.pdf')):
    num = int(''.join(list(filter(str.isdigit, filename))))
    os.rename(filename,f"/Users/maximethomas/Desktop/Test/{num}.pdf")
    file_list.append(num)

print(file_list)