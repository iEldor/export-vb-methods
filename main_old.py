import os

# root path which contains your .vb files
root = 'C:/Users/eldor/Downloads/pytechdegree/export-vb-methods'

code = []
for filename in os.listdir(root):
    if filename.endswith('.vb'):
        file = open(f'{root}/{filename}', 'r')
        lines = file.readlines()
        for line in lines:
            words = line.split()
            for count, word in enumerate(words):
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
                    
file = open('export.csv', 'a')
header = 'File, Method, Validation\n'
file.write(header)

for count, method in enumerate(code):
    file.write(f"{method['file']},{method['method']},{method['validation']}\n")