Attribute VB_Name = "Module 1"
Sub ExportDataToHaEBurp()
    Dim i As Integer
    Dim lastRow As Integer
    Dim myFilePath As String
    Dim myWorkbook As Workbook
    Dim myWorksheet As Worksheet
    
    Set myWorkbook = ActiveWorkbook
    Set myWorksheet = myWorkbook.ActiveSheet
    
    lastRow = myWorksheet.Cells(myWorksheet.Rows.Count, "A").End(xlUp).Row
    
    myWorksheet.Columns("A:A").Select
    Selection.Replace What:="/", Replacement:="\/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    myWorksheet.Columns("A:A").Select
    Selection.Replace What:="{", Replacement:="\{", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    myWorksheet.Columns("A:A").Select
    Selection.Replace What:="}", Replacement:="\}", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    For i = 1 To lastRow
        myWorksheet.Range("B" & i).Value = "  - color: cyan" & vbCrLf _
        & "    engine: nfa" & vbCrLf _
        & "    loaded: true" & vbCrLf _
        & "    name: API " & i & vbCrLf _
        & "    regex: " & myWorksheet.Range("A" & i).Value & vbCrLf _
        & "    scope: request header" & vbCrLf _
        & "    sensitive: false"
    Next i
    
    myWorksheet.Columns("A:A").Select
    Selection.Replace What:="\/", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    myWorksheet.Columns("A:A").Select
    Selection.Replace What:="\{", Replacement:="{", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    myWorksheet.Columns("A:A").Select
    Selection.Replace What:="\}", Replacement:="}", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    MsgBox "Data exported", vbInformation
End Sub
