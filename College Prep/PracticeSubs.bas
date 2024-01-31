Attribute VB_Name = "PracticeSubs"
Option Explicit

Function NameMatch(NameCell As Range, TargetRange As Range) As Range
'Find a student by first and last name. SearchRange is just first names
'Returns the cell for the first name

    Dim c As Range
    
    For Each c In TargetRange
        If NameCell.Value = c.Value And NameCell.Offset(0, 1).Value = c.Offset(0, 1).Value Then
            Set NameMatch = c
            GoTo Break
        End If
    Next

Break:

End Function

Sub CopyAttendance(NameCell As Range, TargetSheet As Worksheet, TargetActivity As String, TargetValue As String, SaveLoad As String)
'Either save attendance from the ActivitiesSheet to the SavedSheet or pull from that's stored in the SavedSheet
    
    Dim TargetCol As Long
    
    'Remove the * if it's there
    TargetActivity = Replace(TargetActivity, "* ", "")
    
    If SaveLoad = "save" Then
        TargetCol = TargetSheet.Range("1:1").Find(TargetActivity, , xlValues, xlWhole).Column
        TargetSheet.Cells(NameCell.Row, TargetCol).Value = TargetValue
    ElseIf SaveLoad = "load" Then
        TargetCol = 1
        TargetSheet.Cells(NameCell.Row, TargetCol).Value = TargetValue
    End If

End Sub

Sub SaveDescription(SaveLoad As String)
'Pulls text from the narrative cell and saves it on a different sheet

    Dim ActivitiesSheet As Worksheet
    Dim DescriptionSheet As Worksheet
    Dim DescriptionRange As Range
    Dim PracticeString As String
    Dim DescriptionString As String
    Dim LRow As Long
    Dim i As Long
    
    Set ActivitiesSheet = Worksheets("Activities Page")
    Set DescriptionSheet = Worksheets("Saved Descriptions")
    Set DescriptionRange = ActivitiesSheet.Range("B3")
    PracticeString = ActivitiesSheet.Range("B1")
    PracticeString = Replace(PracticeString, "* ", "")
    DescriptionString = DescriptionRange.Value
    LRow = DescriptionSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    If SaveLoad = "save" Then GoTo Save
    
    For i = 0 To LRow
        If DescriptionSheet.Range("A1").Offset(i, 0).Value = PracticeString Then
            DescriptionRange.Value = DescriptionSheet.Range("A1").Offset(i, 1).Value
        End If
    Next
    GoTo Footer

Save:
    For i = 0 To LRow
        If DescriptionSheet.Range("A1").Offset(i, 0).Value = PracticeString Then
            DescriptionSheet.Range("A1").Offset(i, 1).Value = DescriptionString
        End If
    Next
    
Footer:
    
End Sub

Sub PopulateSaved()
    'Push students to "Saved Activities" Sheet
    'If we already have students, we only want to push differences and preserve saved attendance
    
    Dim RosterSheet As Worksheet
    Dim SavedSheet As Worksheet
    Dim SavedLRow As Long
    Dim SavedRange As Range
    Dim SavedC As Range
    Dim RosterLRow As Long
    Dim RosterLCol As Long
    Dim RosterRange As Range
    Dim RosterTableStart As Range
    Dim RosterC As Range
    Dim MissingStudents As String
    
    Set RosterSheet = Worksheets("Roster Page")
    Set SavedSheet = Worksheets("Saved Activities")
    Set RosterTableStart = RosterSheet.Range("A6")
    SavedLRow = SavedSheet.Cells(Rows.Count, 1).End(xlUp).Row
    RosterLRow = RosterSheet.Cells(Rows.Count, 2).End(xlUp).Row
    Set RosterRange = RosterSheet.Range(Cells(RosterTableStart.Row + 1, 2).Address, Cells(RosterLRow, 2).Address)
    Set SavedRange = SavedSheet.Range(Cells(2, 1).Address, Cells(SavedLRow, 1).Address)
    
    'If there are no students present, we can just copy and paste
    If CheckTableLength(SavedSheet, SavedSheet.Range("A1")) = False Then
        RosterSheet.Range(Cells(RosterTableStart.Row + 1, 2).Address, Cells(RosterLRow, 3).Address).Copy
        SavedSheet.Range("A2").PasteSpecial xlPasteValues
        
    'Matching names are appended with 1 on both sheets
    Else
        For Each SavedC In SavedRange
            If NameMatch(SavedC, RosterRange) Is Nothing Then
                MissingStudents = MissingStudents & vbCr & SavedC.Value & " " & SavedC.Offset(0, 1).Value
            Else
                Set RosterC = NameMatch(SavedC, RosterRange)
                SavedC.Value = SavedC.Value & "1"
                RosterC.Value = RosterC.Value & "1"
            End If
        Next
    End If

    'Delete students on Saved Activities sheet that are not on the roster
    'First give a warning
    Dim NewStudents As String
    Dim i As Long
    Dim j As Long
    
    If Len(MissingStudents) > 0 Then
        Dim WarningChoice As Long
        Dim WarningString As String
        
        WarningString = " The following students are no longer on your roster: " & vbCr & MissingStudents & vbCr & "Do you wish to remove them?"
        WarningChoice = MsgBox(WarningString, vbQuestion + vbYesNo + vbDefaultButton2)
        If WarningChoice = vbNo Then
            GoTo SkipDelete
        End If
    End If
    
    For i = SavedLRow To 2 Step -1
        If Not InStr(SavedSheet.Cells(i, 1), "1") > 0 Then
            SavedSheet.Cells(i, 1).EntireRow.Delete
        End If
    Next
    
SkipDelete:
    'Put any new students on the Roster to the Saved Activities sheet
    j = 1
    For i = 1 To RosterLRow - RosterTableStart.Row
        If Not InStr(RosterTableStart.Offset(i, 1), "1") > 0 Then
            NewStudents = NewStudents & vbCr & RosterTableStart.Offset(i, 1).Value & " " & RosterTableStart.Offset(i, 2).Value
            SavedSheet.Cells(SavedLRow + j, 1).Value = RosterTableStart.Offset(i, 1).Value
            SavedSheet.Cells(SavedLRow + j, 2).Value = RosterTableStart.Offset(i, 2).Value
            j = j + 1
        End If
    Next

    If Len(NewStudents) > 0 Then
        MsgBox ("Students added: " & NewStudents)
    End If

    'Trim the remaining 1's from first names
    For Each SavedC In SavedRange
        SavedC.Value = Replace(SavedC.Value, "1", "")
    Next

    For Each RosterC In RosterRange
        RosterC.Value = Replace(RosterC.Value, "1", "")
    Next

    'make a table object and add conditional formatting
    Dim RosterTableRange As Range
    
    RosterLCol = RosterSheet.Cells(RosterTableStart.Row, Columns.Count).End(xlToLeft).Column
    Set RosterTableRange = RosterSheet.Range(Cells(RosterTableStart.Row, RosterTableStart.Column).Address, Cells(RosterLRow, RosterLCol).Address)
    RosterSheet.ListObjects.Add(xlSrcRange, RosterTableRange, , xlYes).Name = "AllStudentsTable"
    RosterSheet.ListObjects("AllStudentsTable").ShowTableStyleRowStripes = False
    
    Call TableFormat(RosterSheet.ListObjects("AllStudentsTable"), RosterSheet)
    RosterTableRange.Columns.AutoFit

    'Add Marlett boxes
    Dim BoxRange As Range
    Dim SelectAllRange As Range
    
    Set BoxRange = RosterSheet.Range(Cells(RosterTableStart.Row + 1, 1).Address, Cells(RosterLRow, 1).Address)
    Set SelectAllRange = RosterTableStart.Offset(-1, 0)
    
    Call AddMarlettBox(BoxRange, RosterSheet)
    Call AddSelectAll(SelectAllRange, RosterSheet)
    
Footer:

    'Reprotect
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

Sub CopyRoster(PasteRange As Range)

    Dim ActivitiesSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim LRow As Long
    Dim LCol As Long
    Dim TableRange As Range
    Dim RosterTableStart As Range
    
    Set ActivitiesSheet = Worksheets("Activities Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    
    LRow = RosterSheet.Cells(Rows.Count, 2).End(xlUp).Row
    LCol = RosterSheet.Cells(RosterTableStart.Row, Columns.Count).End(xlToLeft).Column
    
    If LRow = RosterTableStart.Row Then
        MsgBox ("Your roster is empty." & vbCr & _
        "Please paste in your student list")
        Exit Sub
    End If
    
    Set TableRange = RosterSheet.Range(Cells(RosterTableStart.Row, 1).Address, Cells(LRow, LCol).Address)
    TableRange.Copy
    PasteRange.PasteSpecial xlPasteValues
    
End Sub

Sub AutoSaveActivity(ActivityString As String, SaveLoad As String)
'Catpures which students are checked and saves it for reference on a hidden sheet
'Meant for going back and editing activities
'Use "save" and "load" to determine which sheet is updated

    Dim ActivitiesSheet As Worksheet
    Dim SavedSheet As Worksheet
    Dim TableStart As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ActivitiesSheet = Worksheets("Activities Page")
    Set SavedSheet = Worksheets("Saved Activities")
    Set TableStart = ActivitiesSheet.Range("A6")
    
    'Unprotect
    Call UnprotectCheck(SavedSheet)
    
    'Make sure we have students and selected an activity
    If CheckTableLength(ActivitiesSheet, TableStart.Offset(0, 1)) = False Then
        MsgBox ("You have no students added")
        GoTo Footer
    End If
    
    If CheckTableLength(SavedSheet, SavedSheet.Range("A1")) = False Then
        MsgBox ("Something went wrong" & vbCrLf & "Please repull the roster")
        GoTo Footer
    End If
    
    If Len(ActivityString) < 1 Then
        MsgBox ("Please select a practice")
        GoTo Footer
    End If
    
    'Save the attendees for later retrieval
    Dim LRow As Long
    Dim SavedLRow As Long
    Dim i As Long
    Dim IsChecked As String
    Dim MatchRange As Range
    Dim MatchCell As Range
    Dim ActivityCell As Range
    
    LRow = ActivitiesSheet.Cells(Rows.Count, 1).End(xlUp).Row
    SavedLRow = SavedSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Set ActivityCell = SavedSheet.Range("1:1").Find(ActivityString, , xlValues, xlWhole)
    Set MatchRange = SavedSheet.Range(Cells(2, 1).Address, Cells(SavedLRow, 1).Address)
    
    'Change the activity on the Activities sheet since that cell is referenced later
    With ActivitiesSheet.Range("B1")
        .Value = ActivityString
        .WrapText = False
    End With

    If SaveLoad = "save" Then GoTo Save
    
    Set MatchRange = ActivitiesSheet.Range(Cells(TableStart.Row, 2).Address, Cells(LRow, 2).Address)
    For i = 1 To SavedLRow - 1
        IsChecked = ActivityCell.Offset(i, 0)
        Set MatchCell = NameMatch(SavedSheet.Range("A1").Offset(i, 0), MatchRange)
        If MatchCell Is Nothing Then
            MsgBox ("Student " & SavedSheet.Range("A1").Offset(i, 0).Value & " " & SavedSheet.Range("A1").Offset(i, 1) & " can't be found")
            GoTo Footer
        Else
            Call CopyAttendance(MatchCell, ActivitiesSheet, ActivityString, IsChecked, "load")
            Call SaveDescription("load")
        End If
    Next
    GoTo Footer

Save:
    Set MatchRange = SavedSheet.Range(Cells(2, 1).Address, Cells(SavedLRow, 1).Address)
    For i = 1 To LRow - TableStart.Row
        IsChecked = TableStart.Offset(i, 0)
        Set MatchCell = NameMatch(TableStart.Offset(i, 1), MatchRange)
        If MatchCell Is Nothing Then
            MsgBox ("Student " & TableStart.Offset(i, 1).Value & " " & TableStart.Offset(i, 2) & " can't be found")
            GoTo Footer
        Else
            Call CopyAttendance(MatchCell, SavedSheet, ActivityString, IsChecked, "save")
            SaveDescription ("save")
        End If
    Next
 
    'Protect
    Call ResetProtection
 
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub




