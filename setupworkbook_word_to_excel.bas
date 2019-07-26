Attribute VB_Name = "setupworkbook_word_to_excel"
'HOW TO EXPORT MULTIPLE TABLES IN A WORD DOCUMENT INTO SEPARATE EXCEL SHEETS: 'https://stackoverflow.com/questions/4465212/macro-to-export-ms-word-tables-to-excel-sheets
'I have modified the original code found  and commented more

Option Explicit
Sub ImportWordTable()

'setting variables names as either integers, or variant, long or objects
Dim wdDoc As Object
Dim wdFileName As Variant 'variant means you can enter any value, text, numbers
Dim tableNo As Integer 'table number in Word
Dim iRow As Long 'row index in Excel
Dim iCol As Integer 'column index in Excel
Dim resultRow As Long
Dim tableStart As Integer
Dim tableTot As Integer
Dim p, q, y, c As Integer ' Define p and q as integers
Dim x As Variant 'as textbox

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

With wdDoc ' For that specific word file:
    tableNo = wdDoc.Tables.Count ' This counts the number of tables with the document
    tableTot = wdDoc.Tables.Count ' This stores the total number of tables
    If tableNo = 0 Then ' If no table was found, display
        MsgBox "This document contains no tables", _
        vbExclamation, "Import Word Table"
 
    End If

    '-----------------------------------------------------
    'Creating number of empty excel sheets equal to the number of tables counted above
    Worksheets.Add after:=ActiveSheet, Count:=tableTot  ' This means: create as many sheets as the number of tables identifies in the document (Note: Each new workbook comes with one sheet)
    p = Worksheets.Count 'Count the number of worksheets in the workbook and store the number in variable p
    x = InputBox("This Word document contains " & tableNo & " tables. Provide a name pattern for the excel sheets. e.g: if you enter Table, the pattern will be Table1, Table 2, Table3,...", "Enter", "Table")
    y = InputBox("Enter the table number from all the tables detected you want to start extracting from. e.g: If you enter 1, extraction will start from the first table", "Enter", "1")
    
    
     'Beginning of the Loop for all the worksheets created
    For q = 1 To p 'this create a list of numbers from 1 to p
    With Worksheets(q) ' This means we start with the sheet q (e.g: start with sheet 1, then sheet 2,...)
    ' This will rename each table according to a naming pattern that I decided
    .Name = x & q 'I want the sheetname to have the string from input x and the table number
    End With
    Next q 'Start with the next sheet
    '----------------------------------------------------------
    
    c = InputBox("Provide excel sheet number you want to start pasting tables from. e.g: if you enter 1, tables will be pasted starting from the first excel sheet", "Enter", "1")
    
    Worksheets(c).Activate ' Activate the first worksheet to start copying tables
    Range("A:AZ").ClearContents 'Clearing the content of the current worksheet before pasting
    
    For tableStart = y To tableTot 'This is the beginning of the loop. This creates a list of all tables found
        
        ' THE NEXT 3 LINES OF CODE ARE THE HEART OF THIS PROCEDURE: THEY WILL GO AND COPY THE WHOLE TABLE FROM WORD (RANGE) AND PASTE THAT IN THE CURRENT WORKSHEET :
        'BIG BIG HELP : https://stackoverflow.com/questions/50969076/how-to-export-merged-table-from-word-to-excel-using-vba
        
        .Tables(tableStart).Range.Copy
        Range("A2").Activate
        Application.CommandBars.ExecuteMso "PasteSourceFormatting"
            
         
         
         ' The line of code below activates the next excel sheet (I already activated the first worksheet in the beginning), so the table can be pasted there
        Worksheets(ActiveSheet.Index Mod Worksheets.Count + 1).Select      'Solution comes from: https://www.mrexcel.com/forum/excel-questions/25101-how-go-next-worksheet-vba.html
        'Worksheets(ActiveSheet.Index + 1).Select ' This is an alternative #1 to the code right above, to activate the next sheet.
        'ActiveSheet.Next.Activate 'Alternative #2 to activate the next worsheet
        
        'This next line clears all content included in the next sheet before we paste the second table into it
        Range("A:AZ").ClearContents ' Solution found here: https://analysistabs.com/excel-vba/clear-cells-data-range-worksheet/
        
    Next tableStart ' This starts to process the next table

End With
End Sub

' BONUS: HOW TO CREATE MULTIPLE EMPTY EXCEL SHEETS FASTER: https://www.extendoffice.com/documents/excel/2889-excel-create-multiple-sheets-with-same-format.html
'Note: If the program count 50 tables within the document, it's better to create 50 worksheets before running it the second time




