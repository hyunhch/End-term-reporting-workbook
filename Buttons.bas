Attribute VB_Name = "Buttons"
Option Explicit

Sub ReadRoster()
'Read in the roster, make a sortable/filterable table, add Marlett boxes, conditional formatting

    Dim RosterSheet As Worksheet
    Dim SavedSheet As Worksheet
    Dim RosterTableStart As Range
    Dim ColNames() As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    Set SavedSheet = Worksheets("Saved Activities")
    
    'Unprotect
    Call UnprotectCheck(RosterSheet)
    
    'The column headers need to remain unprotected to allow sorting
    'However, the column names and order need to remain constant
    ColNames = Split("Select;First;Last;Ethnicity;Gender;Grade;School;District;Notes", ";")
    Call ResetColumns(RosterSheet, RosterTableStart, ColNames)
    
    'Remove any formatting in the header
    If RosterSheet.AutoFilterMode = True Then
        RosterSheet.AutoFilterMode = False
    End If
    
    'Delete any table objects
    Dim OldTable As ListObject
    
    For Each OldTable In RosterSheet.ListObjects
        OldTable.Unlist
    Next OldTable
    
    'Make sure there are some students added
    If CheckTableLength(RosterSheet, RosterTableStart.Offset(0, 1)) = False Then
        MsgBox ("Please add at least one student.")
        GoTo Footer
    End If
    
    'Populate the Saved Activities sheet
    Call UnprotectCheck(SavedSheet)
    Call PopulateSaved

Footer:
    'Reprotect
    Call ResetProtection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub ResetRosterSheet()
'Clear everything in the Roster Page

    Dim RosterSheet As Worksheet
    Dim RosterTableStart As Range
    Dim LastHeader As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTableStart = RosterSheet.Range("A6")
    
    'Unprotect
    Call UnprotectCheck(RosterSheet)
    
    'Delete everything
    Call ClearSheet(RosterTableStart.Offset(1, 0), 0, RosterSheet)
    
    'Put the column headers back in
    Dim ColNames() As String

    ColNames = Split("Select;First;Last;Ethnicity;Gender;Grade;School;District;Notes", ";")
    Call ResetColumns(RosterSheet, RosterTableStart, ColNames)
    
    'Delete anything past the last column
    Set LastHeader = RosterTableStart.Offset(0, UBound(ColNames))
    Call ClearSheet(LastHeader.Offset(0, 1), 1, RosterSheet)
    
    'Reprotect
    Call ResetProtection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub SelectAll()
'For current sheet. Looks for Marlett font

    Dim FRow As Long
    Dim LRow As Long
    Dim CheckRange As Range
    Dim i As Long
    
    FRow = ActiveSheet.Range("A:A").Find("Select", LookIn:=xlValues).Row
    
    'In case the column name was changed or there is some other problem
    If Not FRow > 0 Then
        MsgBox ("There is a problem with the table." & vbCr & _
            "Please make sure the first column is named ""Select""")
        Exit Sub
    End If
    
    LRow = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    'Check that there is at least row of students
    If Not LRow > FRow Then
        MsgBox ("Please add at least one student to the table.")
        Exit Sub
    End If
    
    Set CheckRange = ActiveSheet.Range(Cells(FRow + 1, 1).Address, Cells(LRow, 1).Address)
    CheckRange.Font.Name = "Marlett"
    
    'Check all if any are blank, uncheck all if none are blank
    'Only apply to visible cells
    If Application.CountIf(CheckRange, "a") = LRow - (FRow) Then
        CheckRange.SpecialCells(xlCellTypeVisible).Value = ""
    Else
        CheckRange.SpecialCells(xlCellTypeVisible).Value = "a"
    End If

End Sub

Sub PullRoster()
'Pull students from the roster page to the cover page

    Dim ActivitiesSheet As Worksheet
    Dim TableStart As Range
    Dim TableRange As Range
    Dim LRow As Long
    Dim LCol As Long
    Dim i As Long
    'Dim SelectAllRange As Range
    Dim SelectButton As Shape
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ActivitiesSheet = Worksheets("Activities Page")
    Set TableStart = ActivitiesSheet.Range("A6")
    
    'Unprotect
    Call UnprotectCheck(ActivitiesSheet)
    
    'Clear contents first
    Call ClearSheet(TableStart, 1, ActivitiesSheet)
    
    'Copy over the roster and verify we have students
    Call CopyRoster(TableStart)
    
    LRow = ActivitiesSheet.Cells(Rows.Count, TableStart.Offset(0, 1).Column).End(xlUp).Row
    LCol = ActivitiesSheet.Cells(TableStart.Row, Columns.Count).End(xlToLeft).Column
    
    If LRow < TableStart.Row + 1 Then
        MsgBox ("There aren't any students here." & vbCr & _
        "Please enter your students on the roster page")
        GoTo Footer
    End If
    
    'Make table object. Unlock the cells in the table to allow for sorting
    Set TableRange = ActivitiesSheet.Range(Cells(TableStart.Row, TableStart.Column), Cells(LRow, LCol))
    ActivitiesSheet.ListObjects.Add(xlSrcRange, TableRange, , xlYes).Name = "StudentTable"
    ActivitiesSheet.ListObjects("StudentTable").ShowTableStyleRowStripes = False
    TableRange.Locked = False

    'Add Marlett Boxes and Select boxes
    Dim BoxRange As Range
    Set BoxRange = ActivitiesSheet.Range(Cells(TableStart.Row + 1, 1).Address, Cells(LRow, 1).Address)

    Call AddMarlettBox(BoxRange, ActivitiesSheet)
    'Call AddSelectAll(SelectAllRange, ActivitiesSheet)

    'Conditional Formatting
    Call TableFormat(ActivitiesSheet.ListObjects("StudentTable"), ActivitiesSheet)
    TableRange.Columns.AutoFit
    ActivitiesSheet.Range("A:A").Columns.AutoFit
    
    'Add in totals to the Report sheet
    Call PullReportTotals
    
Footer:
    'Reprotect
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub RemoveSelected()
'Remove selected rows from a sheet

    Dim DelSheet As Worksheet
    Dim IsChecked As Range
    Dim LRow As Long
    Dim TableStart As Range
    Dim ProtectedSheet As Boolean
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set DelSheet = ActiveSheet
    
    'Unprotect
    Call UnprotectCheck(DelSheet)
    
    'Find where the table starts. This should be the same on every sheet
    Set TableStart = DelSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    
    If TableStart Is Nothing Then
        MsgBox ("Something has gone wrong. Please try on a fresh sheet")
        GoTo Reprotect
    End If
    
    'Make sure we have at least one filled row
    If CheckTableLength(DelSheet, TableStart) = False Then
        MsgBox ("You don't have any students on this page.")
        GoTo Reprotect
    End If
    
    'Loop backward through the rows
    Dim NumChecked As Long
    Dim i As Long
    
    LRow = DelSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = LRow To TableStart.Row + 1 Step -1
        If DelSheet.Cells(i, 1).Value <> "" Then
            DelSheet.Cells(i, 1).EntireRow.Delete
            NumChecked = NumChecked + 1
        End If
    Next i

    If NumChecked < 1 Then
        MsgBox ("You don't have any rows selected.")
    End If
    
Reprotect:
    Call ResetProtection
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SaveActivity(SaveLoad As String)
'Catpures which students are checked and saves it for reference on a hidden sheet
'Meant for going back and editing activities
'Use "save" and "load" to determine which sheet is updated

    Dim ActivitiesSheet As Worksheet
    Dim SavedSheet As Worksheet
    Dim TableStart As Range
    Dim ActivityString As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ActivitiesSheet = Worksheets("Activities Page")
    Set SavedSheet = Worksheets("Saved Activities")
    Set TableStart = ActivitiesSheet.Range("A6")
    ActivityString = ActivitiesSheet.Range("B1").Value
    
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
    
    'Trim off * if it's present
    If InStr(ActivityString, "* ") Then
        ActivityString = Replace(ActivityString, "* ", "")
    End If
    
    'Save the attendees for later retrieval
    Dim LRow As Long
    Dim SavedLRow As Long
    Dim i As Long
    Dim IsChecked As String
    Dim MatchRange As Range
    Dim MatchCell As Range
    Dim ActivityCell As Range
    
    'Trim off * if it's present
    If InStr(ActivityString, "* ") Then
        ActivityString = Replace(ActivityString, "* ", "")
    End If
        
    LRow = ActivitiesSheet.Cells(Rows.Count, 1).End(xlUp).Row
    SavedLRow = SavedSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Set ActivityCell = SavedSheet.Range("1:1").Find(ActivityString, , xlValues, xlWhole)
    Set MatchRange = SavedSheet.Range(Cells(2, 1).Address, Cells(SavedLRow, 1).Address)

    If SaveLoad <> "load" Then GoTo Save
    
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
            Call SaveDescription("save")
        End If
    Next
    
    If SaveLoad = "saveall" Then
        GoTo Footer
    End If

    'Edit the original list on the Ref Tables sheet
    'If at least one student was present, put a "*" in front of the practice, remove it if there are none
    Dim RefSheet As Worksheet
    Dim RefRange As Range
    Dim c As Range
    
    Set RefSheet = Worksheets("Ref Tables")
    Set RefRange = RefSheet.Range("ActivitiesTable")
    Set c = RefRange.Find("*" & ActivityString, , xlValues, xlWhole)

    If AnyChecked(TableStart.Offset(1, 0).Row, LRow, ActivitiesSheet) = True Then
        If InStr(c.Value, "* ") = False Then
            c.Value = "* " + c.Value
        End If
    ElseIf InStr(c.Value, "* ") Then
        c.Value = Replace(c.Value, "* ", "")
    End If
     
    If SaveLoad = "saveall" Then
        GoTo Footer
    End If
    
    MsgBox ("Practice saved")
 
Footer:
    'Protect
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

Sub TabulateCaller()
'Here to call TabulateChecked()

    Dim ReportSheet As Worksheet
    Dim WhichSave As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set ReportSheet = Worksheets("Report Page")

    'Unprotect
    Call UnprotectCheck(ReportSheet)
    
    'This is for tabulating a single practice rather than all of them
    WhichSave = "save"
    
    Call TabulateChecked(WhichSave)
    
Footer:
    'Protect
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub ClearReport()
'Delete everything from the report sheet
    
    Dim ReportSheet As Worksheet
    Dim TableStart As Range 'The grand total column
    Dim TotalRange As Range
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
                           
    Dim DescriptionRange As Range
    Dim DelRange As Range
    Dim LRow As Long
    Dim ClearAll As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ReportSheet = Worksheets("Report Page")
    Set TableStart = ReportSheet.Range("C1")

    'Unprotect
    Call UnprotectCheck(ReportSheet)

    'Make the deletion range. It's discontiguous
    With ReportSheet
        Set TotalRange = .Range("C2:C22")
        Set RaceRange = .Range("E2:L22")
        Set GenderRange = .Range("N2:P22")
        Set GradeRange = .Range("R2:Y22")
                                          
        Set DescriptionRange = .Range("Z3:Z22")
    End With
                                                 
            
    Set DelRange = Union(TotalRange, RaceRange, GenderRange, GradeRange, DescriptionRange)

                                               
 
    
    'Yes/No box
    ClearAll = MsgBox("Are you sure you want to clear all content?" & vbCrLf & "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "")
    
    If ClearAll = vbYes Then
        With DelRange
            .ClearContents
            .FormatConditions.Delete
            .Font.Name = "Calibri"
        End With
    End If
    
    'Reprotect
    Call ResetProtection
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub PullTotalCaller()
'Only exists to allow calling PullReportTotals() on a button press

    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim TableStart As Range
    Dim LRow As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    Set TableStart = RosterSheet.Range("A6")

    'Unprotect
    Call UnprotectCheck(ReportSheet)
    
    'Make sure we have students
    If CheckTableLength(RosterSheet, TableStart) = False Then
        MsgBox ("There are no students on the Roster Page")
        GoTo Footer
    End If
    
    Call PullReportTotals
    
Footer:
    'Reprotect
    Call ResetProtection
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub ClearReportTotals()
'Clears the totals row of the report

    Dim ReportSheet As Worksheet
    Dim TotalRange As Range
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
                           
    Dim StartRange As Range
    Dim EndRange As Range
    Dim DelRange As Range
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ReportSheet = Worksheets("Report Page")

    'Unprotect
    Call UnprotectCheck(ReportSheet)
    
    'Define delete area
    Set StartRange = ReportSheet.Range("1:1").Find("White", , xlValues, xlWhole)
    Set EndRange = ReportSheet.Range("1:1").Find("Other Race", , xlValues, xlWhole)
    Set RaceRange = ReportSheet.Range(StartRange.Address, EndRange.Address).Offset(1, 0)
    
    Set StartRange = ReportSheet.Range("1:1").Find("Female", , xlValues, xlWhole)
    Set EndRange = ReportSheet.Range("1:1").Find("Other Gender", , xlValues, xlWhole)
    Set GenderRange = ReportSheet.Range(StartRange.Address, EndRange.Address).Offset(1, 0)
 
    Set StartRange = ReportSheet.Range("1:1").Find("6", , xlValues, xlWhole)
    Set EndRange = ReportSheet.Range("1:1").Find("Other Grade", , xlValues, xlWhole)
    Set GradeRange = ReportSheet.Range(StartRange.Address, EndRange.Address).Offset(1, 0)
    
                                                                                  
                                                                                    
                                                                                         
    
    Set TotalRange = ReportSheet.Range("1:1").Find("Total", , xlValues, xlWhole).Offset(1, 0)
    
    'Delete
    Set DelRange = Union(RaceRange, GenderRange, GradeRange, TotalRange)
                             
                            
    DelRange.ClearContents
    
    'Reprotect
    Call ResetProtection
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub TabulateAll()
'Go through every practice and tabulate

    Dim ReportSheet As Worksheet
    Dim SavedSheet As Worksheet
    Dim PracticeString As String
    Dim AllPractices As Range
    Dim c As Range
    Dim i As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ReportSheet = Worksheets("Report Page")
    Set SavedSheet = Worksheets("Saved Activities")
    Set AllPractices = SavedSheet.Range("C1:V1")
    
    'Unprotect
    Call UnprotectCheck(ReportSheet)
    
    Call PullReportTotals
    For Each c In AllPractices
        PracticeString = c.Value
        Call AutoSaveActivity(PracticeString, "load")
        Call TabulateChecked("yes")
    Next c
    
    ReportSheet.Activate
    MsgBox ("All practices tabulated")
    
    'Reprotect
    Call ResetProtection
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub SharePointExport()
'reformat the data and export a new spreadsheet to SharePoint. Use a dynamic name with the center name and date
                
    Dim CopyBook As Workbook
    Dim PasteBook As Workbook
    Dim CoverSheet As Worksheet
    Dim NarrativeSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim OtherSheet As Worksheet
    Dim CenterString As String
    Dim NameString As String
    Dim DateString As String
    Dim SubDate As String
    Dim SubTime As String
    Dim SpPath As String
    Dim FileName As String
    Dim i As Long
    Dim NewSheetNames() As String
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'The four sheets we need
    Set CoverSheet = Worksheets("Cover Page")
    Set NarrativeSheet = Worksheets("Narrative Page")
    Set ReportSheet = Worksheets("Report Page")
    Set OtherSheet = Worksheets("Other Activities")

    'Submission information
    With CoverSheet
        NameString = .Range("B3").Value
        DateString = .Range("B4").Value
        CenterString = .Range("B5").Value
    End With
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")
    
    'Make sure information is filled out
    If Len(CenterString) < 1 Or Len(NameString) < 1 Or Len(DateString) < 1 Then
        MsgBox ("Please enter in your name, center, and the date on the Cover Page.")
        GoTo Footer
    End If
    
    'Create a new workbook
    Set CopyBook = ActiveWorkbook
    Set PasteBook = Workbooks.Add
    
    'Names of the new sheets we'll have
    NewSheetNames = Split("Other;Report;Narrative;Cover", ";")
    With PasteBook
        For i = LBound(NewSheetNames) To UBound(NewSheetNames)
            Sheets.Add().Name = NewSheetNames(i)
        Next i
        .Worksheets("Sheet1").Delete
    End With
        
    'Copy over information
    CoverSheet.Range("A1:B5").Copy
    PasteBook.Worksheets("Cover").Range("A1").PasteSpecial
    NarrativeSheet.Cells.Copy
    PasteBook.Worksheets("Narrative").Cells.PasteSpecial
    ReportSheet.Cells.Copy
    PasteBook.Worksheets("Report").Cells.PasteSpecial
    OtherSheet.Cells.Copy
    PasteBook.Worksheets("Other").Cells.PasteSpecial
    
    'The report page has some buttons to remove
    PasteBook.Worksheets("Report").Buttons.Delete

    'Create a file name based on the center and date of submission. The center *must* be filled
    'Path to the folder these will be saved in
    SpPath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Report%20Submissions/"
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    PasteBook.SaveAs FileName:=SpPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    ActiveWorkbook.Close SaveChanges:=False
    
    MsgBox ("Submitted to SharePoint")
    
Footer:
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub LocalSave()
'As SharePointE

    Dim CopyBook As Workbook
    Dim PasteBook As Workbook
    Dim CoverSheet As Worksheet
    Dim NarrativeSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim OtherSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim SavedSheet As Worksheet
    Dim CenterString As String
    Dim NameString As String
    Dim DateString As String
    Dim SubDate As String
    Dim SubTime As String
    Dim LocalPath As String
    Dim FileName As String
    Dim SaveName As String
    Dim i As Long
    Dim NewSheetNames() As String
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'The six sheets we need
    Set CoverSheet = Worksheets("Cover Page")
    Set NarrativeSheet = Worksheets("Narrative Page")
    Set ReportSheet = Worksheets("Report Page")
    Set OtherSheet = Worksheets("Other Activities")
    Set RosterSheet = Worksheets("Roster Page")
    Set SavedSheet = Worksheets("Saved Activities")

    'Submission information
    With CoverSheet
        NameString = .Range("B3").Value
        DateString = .Range("B4").Value
        CenterString = .Range("B5").Value
    End With
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")
    
    'Make sure information is filled out
    If Len(CenterString) < 1 Or Len(NameString) < 1 Or Len(DateString) < 1 Then
        MsgBox ("Please enter in your name, center, and the date on the Cover Page.")
        GoTo Footer
    End If
    
    'Create a new workbook
    Set CopyBook = ActiveWorkbook
    Set PasteBook = Workbooks.Add
    
    'Names of the new sheets we'll have
    NewSheetNames = Split("Attendance;Other;Report;Narrative;Cover", ";")
    With PasteBook
        For i = LBound(NewSheetNames) To UBound(NewSheetNames)
            Sheets.Add().Name = NewSheetNames(i)
        Next i
        .Worksheets("Sheet1").Delete
    End With
    
    'Copy over information
    CoverSheet.Range("A1:B5").Copy
    PasteBook.Worksheets("Cover").Range("A1").PasteSpecial
    NarrativeSheet.Cells.Copy
    PasteBook.Worksheets("Narrative").Cells.PasteSpecial
    ReportSheet.Cells.Copy
    PasteBook.Worksheets("Report").Cells.PasteSpecial
    OtherSheet.Cells.Copy
    PasteBook.Worksheets("Other").Cells.PasteSpecial
    
    'Combine the roster and saved activities. Students will be in the same order on both
    Dim TableStart As Range
    Dim TableRange As Range
    Dim PasteRange As Range
    Dim ReformatRange As Range
    Dim c As Range
    Dim LRow As Long
    Dim LCol As Long
    Dim NewLCol As Long
    
    Set TableStart = RosterSheet.Range("A6")
    LRow = RosterSheet.Cells(Rows.Count, TableStart.Column).End(xlUp).Row
    LCol = RosterSheet.Cells(TableStart.Row, Columns.Count).End(xlToLeft).Column
    Set TableRange = RosterSheet.Range(Cells(TableStart.Row, TableStart.Offset(0, 1).Column).Address, Cells(LRow, LCol).Address)
    Set PasteRange = PasteBook.Worksheets("Attendance").Range("A1")
    
    TableRange.Copy
    PasteRange.PasteSpecial xlPasteValuesAndNumberFormats
    
    Set TableStart = SavedSheet.Range("C1")
    LRow = SavedSheet.Cells(Rows.Count, 1).End(xlUp).Row
    LCol = SavedSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    Set TableRange = SavedSheet.Range(Cells(TableStart.Row, TableStart.Column).Address, Cells(LRow, LCol).Address)
    Set PasteRange = PasteBook.Worksheets("Attendance").Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1)
    
    TableRange.Copy
    PasteRange.PasteSpecial xlPasteValuesAndNumberFormats
    
    'Change the saved activities section into 1's and 0's instead of a's and blanks
    NewLCol = PasteBook.Worksheets("Attendance").Cells(1, Columns.Count).End(xlToLeft).Column
    Set ReformatRange = PasteBook.Worksheets("Attendance").Range(Cells(2, PasteRange.Column).Address, Cells(LRow, NewLCol).Address)
    
    For Each c In ReformatRange
        If c.Value = "a" Then
            c.Value = "1"
        Else
            c.Value = "0"
        End If
    Next c
    
    'Create a file name based on the center and date of submission. The center *must* be filled
    'Path to the folder these will be saved in
    LocalPath = GetLocalPath(ThisWorkbook.path)
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    'For Win and Mac
    If Application.OperatingSystem Like "*Mac*" Then
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            ActiveWorkbook.Close SaveChanges:=False
            GoTo Footer
        End If
        ActiveWorkbook.SaveAs FileName:=LocalPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        ActiveWorkbook.Close SaveChanges:=False
    Else
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            ActiveWorkbook.Close SaveChanges:=False
            GoTo Footer
        End If
        ActiveWorkbook.SaveAs FileName:=SaveName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        ActiveWorkbook.Close SaveChanges:=False
    End If
    
Footer:
    CoverSheet.Activate
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub ClearAllSaved()
'Button to erase all data in the Saved Activities and Saved Descriptions sheets

    Dim SavedSheet As Worksheet
    Dim DescriptionSheet As Worksheet
    Dim RefSheet As Worksheet
    Dim ClearRange As Range
    Dim c As Range
                           
    Dim ClearAll As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set SavedSheet = Worksheets("Saved Activities")
    Set DescriptionSheet = Worksheets("Saved Descriptions")
    Set RefSheet = Worksheets("Ref Tables")
    
    ClearAll = MsgBox("Are you sure you want to clear all content?" & vbCrLf & "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If ClearAll = vbNo Then
        GoTo Footer
    End If
    
    'Unprotect
    Call UnprotectCheck(SavedSheet)
    Call UnprotectCheck(DescriptionSheet)
    
    'Activities sheet. First two columns are first and last name, first row is activity names
    Set ClearRange = SavedSheet.Range("C2")
    Call ClearSheet(ClearRange, 1, SavedSheet)
    
    'Descriptions. Fist column is activity names, everything else can go
    Set ClearRange = DescriptionSheet.Range("B1")
    Call ClearSheet(ClearRange, 1, DescriptionSheet)
    
    'Remove any "*"s from the Ref Tables sheet
    For Each c In RefSheet.Range("ActivitiesTable")
        If InStr(c.Value, "* ") Then
            c.Value = Replace(c.Value, "* ", "")
        End If
    Next c
        
Footer:
    'Reprotect
    Call ResetProtection
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub




