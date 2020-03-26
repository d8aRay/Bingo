# BINGO
Quick script to generate an excel workbook with 5 BINGO boards in separate sheets.
Enjoy!

To automatically highlight, set up the following rule for conditional formatting.
=NOT(ISNA(VLOOKUP(A2,$I:$I,1,FALSE)))
and make sure it applies to 
=$A$2:$E$6
Now, you can track the numbers that have been called in column I, and the board will automagically highlight.
