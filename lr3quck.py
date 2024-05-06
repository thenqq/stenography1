import docx

def take_text(path):
    qwe = []
    for p in path.paragraphs:
        if '\n' in p.text:
            for j in p.text.split('\n'):
                qwe.append(j)
        else:
            qwe.append(p.text)
        p.clear()
    return qwe

doc_path = docx.Document('./cont/2.docx')

text_array = take_text(doc_path)
print(text_array)
print('')
alphabet = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя'
p = '.,/!:;()-—'

y = []
n = []
for i in text_array:
    counter = 0
    for j in i:
        if j in p:
            y.append(i)
            break
        else:
            counter += 1
    if counter == len(i):
        n.append(i)

yy = []
for i in y:
    if len(i) % 2 == 0:
        yy.append(i)
    else:
        n.append(i)
print('+ ', len(yy), ' -> ', yy)
print('- ', len(n), ' -> ', n)
print('')

for i in text_array:
    if i in yy:
        print(i, end='| yes \n')
    elif i in n:
        print(i, end='| no \n')
