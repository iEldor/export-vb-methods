import os

# main project folder
root = 'C:/Users/eldor/Downloads/pytechdegree/export-vb-methods'

# option to extract if/elseif business validations along with method names
include_validations = False
if include_validations:
    export = open('methods_with_validations.csv', 'a'); header = 'File, Method, Validation\n'
else:
    export = open('methods_names_only.csv', 'a'); header = 'File, Method\n'
export.write(header)

# extract and write to file
for dirpath, dirnames, filenames in os.walk(root):
    for filename in filenames:
        if filename.endswith('.vb'):
            filepath = os.path.join(dirpath, filename)
            file = open(filepath, 'r')
            lines = file.readlines()
            for line in lines:
                words = line.split()
                for count, word in enumerate(words):
                    if include_validations:
                        if word in ['Sub', 'Function'] and count != len(words)-1:
                            word = f"{word} {words[count + 1].split('(')[0]}"
                            export.write(f'{filename},{word}, \n')
                        elif word in ['If', 'ElseIf'] and count != len(words)-1:
                            export.write(f"{filename}, ,{' '.join(words)}\n")
                    else:
                        if word in ['Sub', 'Function'] and count != len(words)-1:
                            word = f"{word} {words[count + 1].split('(')[0]}"
                            export.write(f'{filename},{word}\n')