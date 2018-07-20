with open('DMSTotal.txt','r', encoding='utf8') as ft:
    with open('DMSTable.txt', 'r',  encoding='utf8') as fu:
        ftSet = set([])
        fuSet = set([])
        for line in ft:
            uid = int(line.split(',')[0])
            ftSet.update([uid])
        for line in fu:
            uid = int(line.split(',')[0])
            fuSet.update([uid])
        candidate = ftSet.difference(fuSet)

with open('DMSTotal.txt','r',  encoding='utf8') as ft:
    with open('Final.csv', 'w', encoding='utf8') as fout:
        for line in ft:
            uid = int(line.split(',')[0])
            if uid in candidate:
                fout.write(line)
