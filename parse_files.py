import os
import glob

path = 'path to yout directory'

file_list = []
for filename in glob.glob(os.path.join(path, '*.pdf')):
    num = int(''.join(list(filter(str.isdigit, filename))))
    os.rename(filename,f"path to yout directory{num}.pdf")
    file_list.append(num)

print(file_list)
