import os

# main project folder
root = 'C:/Users/eldor/Downloads/pytechdegree/export-vb-methods'

# option to extract if/elseif validations along with method names
include_validations = False

code = []
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
                            code.append({
                                'file': filename
                                ,'method': f'{word} {words[count + 1]}'
                                ,'validation': ' '
                            })
                        elif word in ['If', 'ElseIf'] and count != len(words)-1:
                            code.append({
                                'file': filename
                                ,'method': ' '
                                ,'validation': ' '.join(words)
                            })
                    else:
                        if word in ['Sub', 'Function'] and count != len(words)-1:
                            code.append({
                                'file': filename
                                ,'method': f'{word} {words[count + 1]}'
                            })
                    
if include_validations:
    file = open('methods_with_validations.csv', 'a')
    header = 'File, Method, Validation\n'
    file.write(header)
    for count, method in enumerate(code):
        file.write(f"{method['file']},{method['method']},{method['validation']}\n")
else:
    file = open('methods_names_only.csv', 'a')
    header = 'File, Method\n'
    file.write(header)
    for count, method in enumerate(code):
        file.write(f"{method['file']},{method['method']}\n")