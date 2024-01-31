Attribute VB_Name = "TabulateSubs"
Option Explicit
Option Compare Text

Function DemoTabulate(NumChecked As Long, SearchRange As Range, SearchType As String) As Variant
'Count how many fall into each demographic category and return an array
'Can specify if we're looking for race, gender, or grade

    Dim RaceArray As Variant
    Dim GenderArray As Variant
    Dim GradeArray As Variant
                             
    Dim SearchArray As Variant
    Dim CountArray As Variant
    Dim SearchTerm As String
    Dim SearchHere As Range
    Dim i As Long
    Dim c As Range
    
    RaceArray = Array("White", "Asian", "Black", "Latino", "AIAN", "NHPI", "Mixed", "Other")
    GenderArray = Array("Female", "Male", "Other")
    GradeArray = Array("6", "7", "8", "9", "10", "11", "12", "Other")
                                                                                            
    'What we are searching
    If SearchType = "Race" Then
        ReDim SearchArray(0 To UBound(RaceArray))
        ReDim CountArray(0 To UBound(RaceArray))
        SearchArray = RaceArray
        Set SearchHere = SearchRange.Offset(0, 2)
        
    ElseIf SearchType = "Gender" Then
        ReDim SearchArray(0 To UBound(GenderArray))
        ReDim CountArray(0 To UBound(GenderArray))
        SearchArray = GenderArray
        Set SearchHere = SearchRange.Offset(0, 3)
        
    ElseIf SearchType = "Grade" Then
        ReDim SearchArray(0 To UBound(GradeArray))
        ReDim CountArray(0 To UBound(GradeArray))
        SearchArray = GradeArray
        Set SearchHere = SearchRange.Offset(0, 4)
        GoTo GradeCount
    End If
    
    
    ReDim CountArray(0 To UBound(SearchArray))
   
    'CountIf doesn't work with discontiguous ranges
    For i = 0 To UBound(SearchArray)
        SearchTerm = SearchArray(i)
        For Each c In SearchHere
            If Trim(c.Value) = SearchTerm Then
                CountArray(i) = CountArray(i) + 1
            End If
        Next c
    Next i
    GoTo MissingValues

GradeCount:
    'We can't use Trim() for this
    For i = 0 To UBound(SearchArray)
        SearchTerm = SearchArray(i)
        For Each c In SearchHere
            If c.Value = SearchTerm Then
                CountArray(i) = CountArray(i) + 1
            End If
        Next c
              
    Next i
              
MissingValues:
    'Blank and invalid entries aren't counted above
    Dim Missing As Long
    Dim ArrayIndex As Long
    
    Missing = NumChecked - WorksheetFunction.Sum(CountArray)
    ArrayIndex = UBound(CountArray)
    
    If Missing > 0 Then
        CountArray(ArrayIndex) = CountArray(ArrayIndex) + Missing
    End If
    
    Erase SearchArray
    DemoTabulate = CountArray
    Erase CountArray
    
End Function

Sub PullReportTotals()
'Also called when roster is pulled

    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Unprotect
    Call UnprotectCheck(ReportSheet)

    'Define where the totals will go. It's a discontiguous area
    Dim StartRange As Range
    Dim EndRange As Range
    Dim TotalRange As Range
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
                           
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
    
    'Clear the contents
    TotalRange.ClearContents
    RaceRange.ClearContents
    GenderRange.ClearContents
    GradeRange.ClearContents
    
    'Grab the entire first name column from the Roster Page
    Dim LRow As Long
    Dim TempArray As Variant
    Dim NameRange As Range
    Dim TableStart As Range
    
    Set TableStart = RosterSheet.Range("A6")
    LRow = RosterSheet.Cells(Rows.Count, 2).End(xlUp).Row
    Set NameRange = RosterSheet.Range(Cells(TableStart.Row + 1, 2).Address, Cells(LRow, 2).Address)
    
    'Pass to be tabulated and paste in values
    TempArray = DemoTabulate(LRow - TableStart.Row, NameRange, "Race")
    RaceRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(LRow - TableStart.Row, NameRange, "Gender")
    GenderRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(LRow - TableStart.Row, NameRange, "Grade")
    GradeRange = TempArray
    Erase TempArray
    
    'Total
    TotalRange = LRow - TableStart.Row

End Sub


Sub TabulateChecked(TabAll As String)
'Tabulate demographics for the indicated range and push to the Report Page

    Dim ActivitiesSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim PracticeRange As Range
    Dim PracticeString As String
    Dim DescriptionString As String
    Dim TableStart As Range
    Dim CheckRange As Range
    Dim SearchRange As Range
    Dim LRow As Long
    Dim TotalNumber As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ReportSheet = Worksheets("Report Page")
    Set ActivitiesSheet = Worksheets("Activities Page")
    Set TableStart = ActivitiesSheet.Range("A6")
    DescriptionString = ActivitiesSheet.Range("B3").Value
    
    'Unprotect
    Call UnprotectCheck(ReportSheet)
    
    'Confirm that a practice has been selected. It's fine if the table is empty in this case
    PracticeString = ActivitiesSheet.Range("B1").Value
    
    If Not Len(PracticeString) > 0 Then
        MsgBox ("Please select a practice")
        GoTo Footer
    End If
    
    'Remove any asterisk
    PracticeString = Replace(PracticeString, "* ", "")
    
    'Pull totals for the Report Page
    Call PullReportTotals
                                                 
    'Grab the range to search
    LRow = ActivitiesSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Set CheckRange = ActivitiesSheet.Range(Cells(TableStart.Row + 1, 1).Address, Cells(LRow, 1).Address)
    Set SearchRange = FindChecks(CheckRange)
    TotalNumber = CountChecks(CheckRange)
    
    'Grab the ranges the values will live on the Report Page
    Dim StartRange As Range
    Dim EndRange As Range
    Dim RaceRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim TotalRange As Range
    Dim DescriptionRange As Range
                                 
    Set PracticeRange = ReportSheet.Range("B:B").Find(PracticeString, , xlValues, xlWhole)
    
    'The categories are separated by a subtotal column
    'This should avoid needing to change the code if the columns are changed in the future
    Set StartRange = ReportSheet.Range("1:1").Find("White", , xlValues, xlWhole)
    Set EndRange = ReportSheet.Range("1:1").Find("Other Race", , xlValues, xlWhole)
    Set RaceRange = ReportSheet.Range(StartRange.Address, EndRange.Address).Offset(PracticeRange.Row - 1, 0)
    
    Set StartRange = ReportSheet.Range("1:1").Find("Female", , xlValues, xlWhole)
    Set EndRange = ReportSheet.Range("1:1").Find("Other Gender", , xlValues, xlWhole)
    Set GenderRange = ReportSheet.Range(StartRange.Address, EndRange.Address).Offset(PracticeRange.Row - 1, 0)
 
    Set StartRange = ReportSheet.Range("1:1").Find("6", , xlValues, xlWhole)
    Set EndRange = ReportSheet.Range("1:1").Find("Other Grade", , xlValues, xlWhole)
    Set GradeRange = ReportSheet.Range(StartRange.Address, EndRange.Address).Offset(PracticeRange.Row - 1, 0)

    Set TotalRange = ReportSheet.Range("1:1").Find("Total", , xlValues, xlWhole).Offset(PracticeRange.Row - 1, 0)
    Set DescriptionRange = ReportSheet.Range("1:1").Find("Description", , xlValues, xlWhole).Offset(PracticeRange.Row - 1, 0)
    
    'We can erase anything in the row if there are no selected students
    If TotalNumber = 0 Then
        RaceRange.ClearContents
        GenderRange.ClearContents
        GradeRange.ClearContents
                                
        DescriptionRange.ClearContents
        TotalRange.ClearContents
        GoTo NoStudents
    End If
    
    'Pass to DemoTabulate() which returns a tabulated array
    Dim TempArray As Variant
    Dim AllRange As Range
    
    TempArray = DemoTabulate(TotalNumber, SearchRange.Offset(0, 1), "Race")
    RaceRange = TempArray
    Erase TempArray

    TempArray = DemoTabulate(TotalNumber, SearchRange.Offset(0, 1), "Gender")
    GenderRange = TempArray
    Erase TempArray
    
    TempArray = DemoTabulate(TotalNumber, SearchRange.Offset(0, 1), "Grade")
    GradeRange = TempArray
    Erase TempArray

    'Input student total and description
    TotalRange.Value = TotalNumber
    DescriptionRange.Value = DescriptionString
    
    'Save the practice
    If TabAll = "yes" Then
        Call SaveActivity("saveall")
    Else
        Call SaveActivity("save")
    End If
    
NoStudents:
    'Create a range to pass for adding conditional formatting
    Set AllRange = Union(RaceRange, GenderRange, GradeRange, TotalRange)
    Call FormatGreaterThan(AllRange, ReportSheet.Range("C2"))
    
    ReportSheet.Activate
    
Footer:
    'Protect
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

