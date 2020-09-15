Attribute VB_Name = "InsertRowsOrColumns"
'INSERT MULTIPLE ROWS INTO TABLE

'==================================================================================
'subInsertRows Sub
'----------------------------------------------------------------------------------
'Purpose:   Allows you to easily insert multiple rows at once without copying and
'               pasting
'
'Author:    Rachel J Arthur
'
'Notes:     Particularly useful when working with tables
'           Inserts rows above
'
'----------------------------------------------------------------------------------
'Parameters
'----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------
'Revision History
'----------------------------------------------------------------------------------
'
'Version 1.0.0      01/05/2019      RJA     Initial release
'
'----------------------------------------------------------------------------------
Public Sub subInsertRows()

    Dim intNumberOfRows As Integer
    
    'Get number of rows to insert
    intNumberOfRows = InputBox("How many rows should we insert?", "Number of Rows", 1)
    
    'Insert that many rows at the current active cell
    For i = 1 To intNumberOfRows
        ActiveCell.EntireRow.Insert Shift:=xlShiftDown
    Next i

End Sub

'==================================================================================
'subInsertColumns Sub
'----------------------------------------------------------------------------------
'Purpose:   Allows you to easily insert multiple columns at once without copying and
'               pasting
'Author:    Rachel J Arthur
'
'Notes:     Particularly useful when working with tables
'           Inserts columns to the left
'
'----------------------------------------------------------------------------------
'Parameters
'----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------
'Revision History
'----------------------------------------------------------------------------------
'
'Version 1.0.0      01/05/2019      RJA     Initial release
'
'----------------------------------------------------------------------------------
Public Sub subInsertColumns()

    Dim intNumberOfColumns As Integer
    
    'Get number of columns to insert
    intNumberOfColumns = InputBox("How many columns should we insert?", "Number of Columns", 1)
    
    'Insert that many columns at the current active cell
    For i = 1 To intNumberOfColumns
        ActiveCell.EntireColumn.Insert Shift:=xlShiftRight
    Next i

End Sub