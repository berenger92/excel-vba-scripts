Attribute VB_Name = "Rename_sheets_pattern_method2"
' This explains how to create multiple excel worksheets at once in Excel.
' Most of our words documents have many tables and each of them should be pasted on a separate worksheet
' I found a helpful code here that I modified to create and rename every worksheet on the fly: https://www.youtube.com/watch?v=MfO1p_ErJfk
Option Base 1
Sub AddMultiSheetswithNames()
    Dim p, q As Integer ' Define p and q as integers
    'Dim Sheetname As String
    
    Worksheets.Add after:=Sheet1, Count:=60 ' This means: create 60 other sheets after the first sheet (Note: Each new workbook comes with one sheet)
    p = Worksheets.Count 'Count the number of worksheets in the workbook and store the number in variable p
    
    'Beginning of the Loop for all the worksheets created
    For q = 1 To p 'this create a list of numbers from 1 to p
    With Worksheets(q) ' This means we start with the sheet q (e.g: start with sheet 1, then sheet 2,...)
    
    ' This will rename each table according to a naming pattern that I decided
    .Name = "Table" & q 'I want the sheetname to have the string Table and the table number
    End With
    Next q 'Start with the next sheet
    
End Sub 'To close operation


