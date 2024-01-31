Attribute VB_Name = "MarlettBoxSubs"
Option Explicit

Sub AddMarlettBox(BoxHere As Range, TargetSheet As Worksheet)
'Doing this instead of actual checkboxes to deal with sorting issues
'This only changes the font of a range to Marlett

    With BoxHere
        .Font.Name = "Marlett"
        .Value = ""
    End With

End Sub

Sub AddSelectAll(BoxHere As Range, TargetSheet As Worksheet)
'Insert a button

    Dim NewButton As Button
    
    Set NewButton = TargetSheet.Buttons.Add(BoxHere.Left, BoxHere.Top, _
        BoxHere.Width, BoxHere.Height)
    
    With NewButton
        .OnAction = "SelectAll"
        .Caption = "Select All"
    End With

End Sub

Function AnyChecked(StartRow As Long, StopRow As Long, TargetSheet As Worksheet) As Boolean
'Check to see if any students have been checked

    Dim CheckRange As Range
    Dim CheckCell As Range
    
    AnyChecked = False
    Set CheckRange = TargetSheet.Range(Cells(StartRow, 1).Address, Cells(StopRow, 1).Address)
    
    For Each CheckCell In CheckRange
        If CheckCell.Value = "a" Then
            AnyChecked = True
            Exit Function
        End If
    Next CheckCell
    
End Function

Function FindChecks(TargetRange As Range) As Range
'Returns a range that contains all cells with a checkmark

    Dim CheckedRange As Range
    Dim c As Range
    
    For Each c In TargetRange
        If c.Value <> "" Then
            If Not CheckedRange Is Nothing Then
                Set CheckedRange = Union(CheckedRange, c)
            Else
                Set CheckedRange = c
            End If
        End If
    Next c
    
    Set FindChecks = CheckedRange

End Function

Function CountChecks(TargetRange As Range) As Long
'Returns a range that contains all cells with a checkmark

    Dim CheckedRange As Range
    Dim c As Range
    
    CountChecks = 0
    
    For Each c In TargetRange
        If c.Value <> "" Then
            CountChecks = CountChecks + 1
        End If
    Next c
    
End Function
