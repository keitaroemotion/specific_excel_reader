#!/usr/bin/env python3

# -*- coding: utf-8 -*-

import csv
import matplotlib.pyplot as plt
import openpyxl
import re
import sys
from openpyxl        import Workbook
from openpyxl.styles import Alignment, Border, Side

if len(sys.argv) < 2:
    print("\nyou need to set excel path.\n")
    sys.exit()

xlsx_path = sys.argv[1]

workbook  = openpyxl.load_workbook(xlsx_path) 
sheet     = workbook.active 

answers = {}
for question_index in range(1, 16):
    answers[question_index] = []
    for student in range(2, 15):
        cell = sheet.cell(row = student, column = question_index)
        answers[question_index].append(cell.value)

print(answers)
print("")

def makepiechart(labels, index, _flatten=True):
    explode = [0 for x in labels]
    colors  = ['blue', 'red']  
    if len(labels) != 2:
        colors = ['white', 'yellow', 'pink', 'blue', 'red', 'black', 'green'][0:len(labels)]

    fields = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'][0:len(labels)]

    _ans = answers[index]

    if _flatten == True:
       _ans = to_alphabet(answers, index)

    print(_ans)

    sizes  = [
                 select(_ans, x) \
                 for x
                 in  fields 
             ]

    patches, texts = plt.pie(
                         sizes,
                         labels     = labels,
                         colors     = colors,
                         shadow     = True,
                         startangle = 90
                     )

    plt.legend(patches, labels, loc="best")
    plt.axis('equal')
    plt.tight_layout()
    plt.show()


def select(xs, value):
    return len([x for x in xs if x == value])

def flatten(lst):
    for el in lst:
        if isinstance(el, list): 
            yield from flatten(el)
        else:
            yield el

def to_alphabet(answers, i):
    return list(
               flatten(
                   [
                       [
                           [
                               [char for char in x]
                               for x
                               in  re.search("([A-Z])+", x).group(0)
                           ]
                           for x
                           in answers[i]
                       ]
                   ]
               )
           )

#
# Question1: Gender ratio
#
makepiechart(
    ['Male', 'Female'],
    1
)
  
makepiechart(
    [
        '<30',
        '30-40',
        '40-50',
        '50-60',
        '60-70',
        '70-80',
        '>80',
    ], 
    2
)

makepiechart(
    [
        'Elementary School',
        'Junior High school',
        'High school',
        'College/University',
        'Master',
        'PhD',
    ], 
    3
)

makepiechart(
    [
        'Less than a year',
        '1-3 years',
        '3-5 years',
        '5-10 years',
        'More than 10 years',
        'Else'
    ], 
    4
)

makepiechart(
    [
        'Less than a year',
        '1-3 years',
        '3-5 years',
        '5-10 years',
        'More than 10 years',
        'Else'
    ], 
    5
)

makepiechart(
    [
        'Less than a year',
        '1-3 years',
        '3-5 years',
        '5-10 years',
        'More than 10 years',
        'Else'
    ], 
    5
)

#
# 6. The reason why decided to learn Chinese
#
makepiechart(
    [
        'Job',
        'Travel',
        'Curitosity',
        'Exam',
        'Else'
    ], 
    6,
    True
)

#
# 6. The reason why decided to learn Chinese
#
makepiechart(
    [
        'Job',
        'Travel',
        'Curitosity',
        'Exam',
        'Else'
    ], 
    6,
    True
)

#
# 7. Favorite study method of Chinese learning
#
makepiechart(
    [
        'One-by-one tutorial by tutor',
        'Group study(2-10)',
        'Mass tutorial(more than15)',
        'Else'
    ], 
    7,
    True
)

#
# 8. You want to continue learning at XXXXX?
#
makepiechart(
    [
        'Yes',
        'No'
    ], 
    8,
    True
)

#
# 9. You want to learn the other language?
#
makepiechart(
    [
        'Yes',
        'No'
    ], 
    9,
    True
)

#
# 10. Which event do you want to join?
#
makepiechart(
    [
        'Nomikai',
        'Climing/Sports',
        'Nokai',
        'Travel',
        'Karaoke',
        'Else',
        'All',
    ], 
    10,
    True
)

#
# 11. Do you like your tutor?
#
makepiechart(
    [
        '1',
        '2',
        '3',
        '4',
        '5'
    ], 
    11
)

makepiechart(
    [
        '1',
        '2',
        '3',
        '4',
        '5'
    ], 
    12
)

makepiechart(
    [
        '1',
        '2',
        '3',
        '4',
        '5'
    ], 
    13
)

makepiechart(
    [
        '1',
        '2',
        '3',
        '4',
        '5'
    ], 
    14
)

makepiechart(
    [
        '1',
        '2',
        '3',
        '4',
        '5'
    ], 
    15
)
