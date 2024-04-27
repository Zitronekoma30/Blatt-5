from docx import Document
from docx.shared import Pt
import hashlib
import itertools
import random


def hash_file(filename):
    hasher = hashlib.md5()
    
    with open(filename, 'rb') as file:
        chunk = 0
        while chunk != b'':
            chunk = file.read(8192)
            hasher.update(chunk)
    
    return hasher.hexdigest()

def modify_docx(file_path):
    '''generates a random combination of font size, bold, italic, and underline for each paragraph in the document'''
    doc = Document(file_path)
    changes = []
    for paragraph in doc.paragraphs:
        modif = []
        modif.append(random.randint(1, 61))
        modif.append(random.choice([False, True]))
        modif.append(random.choice([False, True]))
        modif.append(random.choice([False, True]))
        
        for run in paragraph.runs:
            run.font.size = Pt(modif[0])
            
            run.font.bold = modif[1]
            run.font.italic = modif[2]
            run.font.underline = modif[3]
        changes.append(modif)

    doc.save(file_path)
    return changes

def generate_hash_change_list(file_path, i):
    hash_changes = []
    for _ in range(i):
        changes = modify_docx(file_path)
        hash = hash_file(file_path)
        hash_changes.append((hash, changes))
    
    return hash_changes

def birthday_attack(file1, file2, tries):
    hash_changes1 = generate_hash_change_list(file1, tries)
    hash_changes2 = generate_hash_change_list(file2, tries)

    for i in range(tries):
        for j in range(tries):
            if hash_changes1[i][0] == hash_changes2[j][0]:
                print(f'Collision found')
                return hash_changes1[i][1], hash_changes2[j][1]
    print('No collision found')

birthday_attack('fair.docx', 'unfair.docx', 100000)