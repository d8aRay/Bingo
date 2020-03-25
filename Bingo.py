# -*- coding: utf-8 -*-
"""
Created on Wed Mar 25 17:52:41 2020

@author: ray.edwards
"""

import random
import pandas as pd

ranges = {'B': (1,15)
         , 'I': (16,30)
         , 'N': (31,45)
         , 'G': (46,60)
         , 'O': (61,75)}

writer = pd.ExcelWriter('bingo.xlsx', engine='xlsxwriter')
for i in range(1,6):
    board = {'B': []
         , 'I': []
         , 'N': []
         , 'G': []
         , 'O': []}

    for letter, value in board.items():
        while len(board[letter])<5:
            board[letter].append(random.randint(ranges[letter][0], ranges[letter][1] ))
            board[letter] = sorted(set(board[letter]))
    board['N'][2] = 'FREE'
    
    df = pd.DataFrame(board)
    df.to_excel(writer, sheet_name='Sheet'+str(i), index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()