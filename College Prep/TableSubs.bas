Attribute VB_Name = "TableSubs"
Option Explicit

Sub TableFormat(NewTable As ListObject, TargetSheet As Worksheet)
  
    'Blanks
    With NewTable.DataBodyRange
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlBlanksCondition
    End With
    'Clear from the first column
    NewTable.ListColumns(1).DataBodyRange.FormatConditions.Delete
    
    With NewTable.DataBodyRange.FormatConditions(1)
        .StopIfTrue = False
        .Interior.Color = 49407
    End With
    
    'Demographics
    Dim RaceSource As String
    Dim GenderSource As String
    Dim GradeSource As String
    Dim RefSheet As Worksheet
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim c As Range
    Dim FormulaString1 As String
    Dim FormulaString2 As String
    
    Set RefSheet = Worksheets("Ref Tables")
    RaceSource = RefSheet.ListObjects("EthnicityTable").DataBodyRange.Columns(1).Address
    GenderSource = RefSheet.ListObjects("GenderTable").DataBodyRange.Columns(1).Address
    GradeSource = RefSheet.ListObjects("GradeTable").DataBodyRange.Columns(1).Address

    Set RaceRange = NewTable.ListColumns("Ethnicity").DataBodyRange
    Set GenderRange = NewTable.ListColumns("Gender").DataBodyRange
    Set GradeRange = NewTable.ListColumns("Grade").DataBodyRange
    
    For Each c In RaceRange
        FormulaString1 = "=AND(ISERROR(MATCH(TRIM(" + c.Address + ")," + "'" + "Ref Tables" + "'" + "!"
        FormulaString2 = ",0)), NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & RaceSource & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    For Each c In GenderRange
        FormulaString1 = "=AND(ISERROR(MATCH(TRIM(" + c.Address + ")," + "'" + "Ref Tables" + "'" + "!"
        FormulaString2 = ",0)), NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & GenderSource & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c

    'Add VALUE() since we're looking at numbers
    For Each c In GradeRange
        FormulaString1 = "=AND(ISERROR(MATCH(VALUE(TRIM(" + c.Address + "))," + "'" + "Ref Tables" + "'" + "!"
        FormulaString2 = ",0)), NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & GradeSource & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
End Sub

Function CheckTableLength(CheckSheet As Worksheet, CheckStart As Range) As Boolean
'Small sub to make sure there's at least one student

    Dim HaveStudent As Boolean
    Dim i As Long
    
    HaveStudent = False
    i = CheckSheet.Cells(Rows.Count, CheckStart.Column).End(xlUp).Row
    If i > CheckStart.Row Then
        HaveStudent = True
    End If
    
    CheckTableLength = HaveStudent

End Function

Sub ResetColumns(TargetSheet As Worksheet, TargetCell As Range, TargetNames As Variant)
'The default column names
    Dim HeaderRange As Range
    
    Set HeaderRange = TargetSheet.Range(Cells(TargetCell.Row, TargetCell.Column).Address, Cells(TargetCell.Row, UBound(TargetNames) + 1).Address)
    HeaderRange.Value = TargetNames
End Sub

Sub FormatGreaterThan(TargetRange As Range, CompareRange As Range)
'Conditional formatting to flag values that are larger than the total number of students in a category
'I don't think we need this and can just do the conditional formatting by hand
    
    Dim ReportSheet As Worksheet
    Dim c As Range
    Dim FormulaString As String
    
    Set ReportSheet = Worksheets("Report Page")
    
    'This will fail if the sheet is protected
    Call UnprotectCheck(ReportSheet)
    
    For Each c In TargetRange
        FormulaString = "=" + ReportSheet.Cells(CompareRange.Row, c.Column).Address + "<" + c.Address
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString
        With c.FormatConditions(1)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c

End Sub

