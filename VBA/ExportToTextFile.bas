Attribute VB_name ="Module 1"
Sub ExportToFile()
    Dim i As Integer
    Dim lastRow As Integer
    Dim myFSO As Object
    Dim myFile As Object
    Dim myPath As String
    Dim myWorkbook As Workbook
    Dim myWorksheet As Worksheet
    
    Set myWorkbook = ActiveWorkbook
    Set myWorksheet = myWorkbook.ActiveSheet
    
    lastRow = myWorksheet.Cells(myWorksheet.Rows.Count, "A").End(xlUp).Row
    
    myFilePath = Application.ActiveWorkbook.Path & "\Output.txt"
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    Set myFile = myFSO.CreateTextFile(myFilePath, True)
    
    'set the last row to the last non-empty row in column A
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop through each row in column A and write to the text file
    For i = 1 To lastRow
        myFile.WriteLine "Item no: " & i
        myFile.WriteLine "Purpose: Party"
        myFile.WriteLine "Color: White"
        myFile.WriteLine "Item: " & Cells(i, "A").Value
    Next i
    
    'close the text file
    myFile.Close
    
    'display a message box when the code is complete
    MsgBox "Text file created successfully!"
End Sub
