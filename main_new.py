import os
import xlwt

# main project folder
project = 'ProjectName'
root = 'Path'


def add_sheet(book, sheet_name, headers, header_row):
    sheet = book.add_sheet(sheet_name)
    for col in range(len(headers)):
        sheet.write(0, col, headers[col])
    return sheet, header_row

headers1 = ['File Name', 'Method Name']
headers2 = ['File Name', 'Method Name', 'Validation']

book = xlwt.Workbook()
methods_only, rownum1 = add_sheet(book, 'methods_only', headers1, 1)
methods_and_validations, rownum2 = add_sheet(book, 'methods_and_validations', headers2, 1)


for dirpath, dirnames, filenames in os.walk(root):
    for filename in filenames:
        if filename.endswith('.vb') and 'designer' not in filename.lower():
            filepath = os.path.join(dirpath, filename)
            file = open(filepath, 'r', encoding = 'unicode_escape')
            lines = file.readlines()
            for line in lines:
                words = line.split()
                for count, word in enumerate(words):
                    if word in ['Sub', 'Function'] and count != len(words)-1:
                        word = f"{word} {words[count + 1].split('(')[0]}"
                        methods_only.write(rownum1, 0, filename)
                        methods_only.write(rownum1, 1, word)
                        methods_and_validations.write(rownum2, 0, filename)
                        methods_and_validations.write(rownum2, 1, word)
                        rownum1 += 1; rownum2 += 1;
                    elif word in ['If', 'ElseIf'] and count != len(words)-1:
                        methods_and_validations.write(rownum2, 0, filename)
                        methods_and_validations.write(rownum2, 2, ' '.join(words))
                        rownum2 += 1
book.save(f'{project}_conversion_check.xls')
