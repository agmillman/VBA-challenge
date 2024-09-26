# VBA-challenge
Module 2 Challenge

All code is original.  

The following references were used in the development of the code.

Find last row with data: LastRecord = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
https://stackoverflow.com/questions/38882321/better-way-to-find-last-used-row

Unique Function: Tickers = WorksheetFunction.Unique(Range("A2:A" & LastRecord))
https://www.mrexcel.com/board/threads/vba-how-to-use-worksheetfunction-unique.1176348/

Format as Accounting: Cells(x + 1, 10).NumberFormat = "_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* ""-""??_);_(@_)"
https://answers.microsoft.com/en-us/msoffice/forum/all/excel-vba-accounting-format/69b68de4-3750-4a1f-9d41-b70ecd5935fb
