Attribute VB_Name = "SheetSubs"

Option Explicit

Sub UnprotectCheck(TargetSheet As Worksheet)
'Checks if a sheet is protected and unprotects
'Used to avoid trying to unprotect an already unprotected sheet

    If TargetSheet.ProtectContents = True Then
        TargetSheet.Unprotect
    End If

End Sub

Sub ResetProtection()
'Reset all sheet protections
    
    Dim RosterSheet As Worksheet
    Dim ActivitiesSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim CoverSheet As Worksheet
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ActivitiesSheet = Worksheets("Activities Page")
    Set ReportSheet = Worksheets("Report Page")
    Set CoverSheet = Worksheets("Cover Page")

    RosterSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    ActivitiesSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True
    ReportSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    CoverSheet.Protect , userinterfaceonly:=True

    'Lock/Unlock areas
    CoverSheet.Range("B3:B5").Locked = False
    
    RosterSheet.Cells.Locked = False
    RosterSheet.Range("A1:A5").EntireRow.Locked = True
    
    ActivitiesSheet.Cells.Locked = False
    ActivitiesSheet.Range("A1:A5").EntireRow.Locked = True
    ActivitiesSheet.Range("B1").Locked = False 'Name of practice
    ActivitiesSheet.Range("B3").Locked = False 'Practice description
    
    ReportSheet.Cells.Locked = True 'Lock the entire page
End Sub

Sub ClearSheet(DelStart As Range, Repull As Long, TargetSheet As Worksheet)
'Repull = 1 avoids warning message

    Dim DelRange As Range
    Dim ClearAll As Long
    Dim OldTable As ListObject

    Set DelRange = TargetSheet.Range(Cells(DelStart.Row, DelStart.Column).Address, Cells(TargetSheet.Rows.Count, TargetSheet.Columns.Count).Address)
    
    If Repull <> 1 Then
            ClearAll = MsgBox("Are you sure you want to clear all content?" & vbCrLf & "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "")
    Else
        ClearAll = vbYes
    End If
    
    If ClearAll = vbYes Then
        For Each OldTable In TargetSheet.ListObjects
            OldTable.Unlist
        Next OldTable
        
        With DelRange
            .ClearContents
            .ClearFormats
            .Validation.Delete
        End With
    End If
    
End Sub

Function NameDateCenter() As Boolean
'Make sure activity info is filled out

    Dim CoverSheet As Worksheet
    Dim NameString As String
    Dim DateString As String
    
    Set CoverSheet = Worksheets("Cover Page")
    NameString = CoverSheet.Range("B1").Value
    DateString = CoverSheeet.Range("B3").Value
    
    If Len(NameString) < 1 Or Len(DateString) < 1 Then
        MsgBox ("Please fill out your name, date, and center on the Cover Page")
        NameDateCenter = False
    Else
        NameDateCenter = True
    End If

End Function



