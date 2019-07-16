Attribute VB_Name = "Modifiedcode_word_to_excel"
'HOW TO EXPORT MULTIPLE TABLES IN A WORD DOCUMENT INTO SEPARATE EXCEL SHEETS: 'https://stackoverflow.com/questions/4465212/macro-to-export-ms-word-tables-to-excel-sheets
'I have modified the original code found  and commented more

Option Explicit
Sub ImportWordTable()

Dim wdDoc As Object
Dim wdFileName As Variant
Dim tableNo As Integer 'table number in Word
Dim iRow As Long 'row index in Excel
Dim iCol As Integer 'column index in Excel
Dim resultRow As Long
Dim tableStart As Integer
Dim tableTot As Integer

' the next two lines of code are to make the file run faster
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

On Error Resume Next ' This allows the program to continue to the next table even if it finds an error on one table

'ActiveSheet.Range("A:AZ").ClearContents
 Worksheets(1).Activate ' Activate the first worksheet
 
wdFileName = Application.GetOpenFilename("Word files (*.docx),*.docx", , _
"Browse for file containing table to be imported")

If wdFileName = False Then Exit Sub '(user cancelled import file browser)

Set wdDoc = GetObject(wdFileName) 'open Word file

With wdDoc
    tableNo = wdDoc.tables.Count ' This counts the number of tables with the document
    tableTot = wdDoc.tables.Count ' This stores the total number of tables
    If tableNo = 0 Then ' If no table was found
        MsgBox "This document contains no tables", _
        vbExclamation, "Import Word Table"
    ElseIf tableNo > 1 Then ' If more than 1 table was found
        tableNo = InputBox("This Word document contains " & tableNo & " tables." & vbCrLf & _
        "Enter the table to start from", "Import Word Table", "1") ' Display a box asking to load the table number to start from
    End If

    resultRow = 2 'this signals where the row number on which the first table to paste will start

    For tableStart = 1 To tableTot '1 To tableTot : This is the beginning of the loop. This creates a list of all tables founds
                    'You can specify any range of tables. From 1 To 9
        With .tables(tableStart) 'This goes through each table one at a time
        
            'copy cell contents from Word table cells to Excel cells
            For iRow = 1 To .Rows.Count
                For iCol = 1 To .Columns.Count
                    Cells(resultRow, iCol) = WorksheetFunction.Clean(.cell(iRow, iCol).Range.Text)
                Next iCol
                resultRow = resultRow + 1 ' this indicates that it pastes the next result on the next row
            Next iRow
            
         ' The line of code below activates the next excel sheet (I already activated the first worksheet in the beginning), so the table can be pasted there
        Worksheets(ActiveSheet.Index Mod Worksheets.Count + 1).Select      'Solution comes from: https://www.mrexcel.com/forum/excel-questions/25101-how-go-next-worksheet-vba.html
        'Worksheets(ActiveSheet.Index + 1).Select ' This is an alternative #1 to the code right above, to activate the next sheet.
        'ActiveSheet.Next.Activate 'Alternative #2 to activate the next worsheet
        
        'This next line clears all content included in the next sheet before we paste the second table into it
        Range("A:AZ").ClearContents ' Solution found here: https://analysistabs.com/excel-vba/clear-cells-data-range-worksheet/
        
        End With 'This ends the process of just one table
        resultRow = 2    'I reset the rowresult here because I want the next table to start at the beginning of the next worksheet
        'if I replace by resultRow= resultRow+1, the next table will start on the row the previous one finished 'and if all tables are copied within same worksheet, the +1 allows to skip one line before pastin the next table
    Next tableStart ' This starts to process the next table

End With

' THE NEXT TWO LINES OF CODE ARE DESIGNED TO MAKE THE FILE run faster
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

' BONUS: HOW TO CREATE MULTIPLE EMPTY EXCEL SHEETS FASTER: https://www.extendoffice.com/documents/excel/2889-excel-create-multiple-sheets-with-same-format.html
'Note: If the program count 50 tables within the document, it's better to create 50 worksheets before running it the second time
