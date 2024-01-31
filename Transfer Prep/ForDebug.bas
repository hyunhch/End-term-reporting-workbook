Attribute VB_Name = "ForDebug"
Sub ScreenUpdating()

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub CoverPlaceButtons()
'To place some control buttons inside cells

    Dim CoverSheet As Worksheet
    Dim MyButton As Button
    Dim ButtonRange As Range
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'Local Save button
    Set ButtonRange = CoverSheet.Range("D2:E3")
    Set MyButton = CoverSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "LocalSave"
        .Caption = "Save Copy"
    End With
    
    'Export button
    Set ButtonRange = CoverSheet.Range("D4:E5")
    Set MyButton = CoverSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "SharePointExport"
        .Caption = "Submit Report"
    End With

End Sub

Sub RosterPlaceButtons()

    Dim RosterSheet As Worksheet
    Dim MyButton As Button
    Dim ButtonRange As Range
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Read roster button
    Set ButtonRange = RosterSheet.Range("A2:B3")
    Set MyButton = RosterSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "ReadRoster"
        .Caption = "Read Roster"
    End With
    
    'Clear Roster button
    Set ButtonRange = RosterSheet.Range("C2:D3")
    Set MyButton = RosterSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "RosterSheetClear"
        .Caption = "Clear Roster"
    End With

    'Delete Row button
    Set ButtonRange = RosterSheet.Range("B5:C5")
    Set MyButton = RosterSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "RemoveSelected"
        .Caption = "Delete Row"
    End With

End Sub

Sub ReportPlaceButtons()

    Dim ReportSheet As Worksheet
    Dim MyButton As Button
    Dim ButtonRange As Range
    
    Set ReportSheet = Worksheets("Report Page")

    'Clear entire report
    Set ButtonRange = ReportSheet.Range("A23")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "ClearReport"
        .Caption = "Clear Report"
    End With
    
    'Pull totals
    Set ButtonRange = ReportSheet.Range("A25")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "PullTotalCaller"
        .Caption = "Pull Totals"
    End With
    
    'Clear totals
    Set ButtonRange = ReportSheet.Range("A26")
    Set MyButton = ReportSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "ClearReportTotals"
        .Caption = "Clear Totals"
    End With

End Sub

Sub ActivitiesPlaceButtons()

    Dim ActivitiesSheet As Worksheet
    Dim PracticeRange As Range
    Dim MyButton As Button
    Dim ButtonRange As Range
    
    Set ActivitiesSheet = Worksheets("Activities Page")
    Set PracticeRange = ActivitiesSheet.Range("B1")

    'Pull Roster button
    Set ButtonRange = ActivitiesSheet.Range("C5:D5")
    Set MyButton = ActivitiesSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "PullRoster"
        .Caption = "Pull Roster"
    End With
    
    'Select all button
    Set ButtonRange = ActivitiesSheet.Range("A5")
    Set MyButton = ActivitiesSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "SelectAll"
        .Caption = "Select All"
    End With
    
    'Delete row
    Set ButtonRange = ActivitiesSheet.Range("A4")
    Set MyButton = ActivitiesSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "RemoveSelected"
        .Caption = "Delete Row"
    End With

    'Save activity
    Set ButtonRange = ActivitiesSheet.Range("E4:F4")
    Set MyButton = ActivitiesSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "'SaveActivity " & """save""" & "'"
        .Caption = "Save Practice"
    End With

    
    'Load activity
    Set ButtonRange = ActivitiesSheet.Range("E5:F5")
    Set MyButton = ActivitiesSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "'SaveActivity " & """load""" & "'"
        .Caption = "Load Practice"
    End With
    
    'Tabulate Activity
    Set ButtonRange = ActivitiesSheet.Range("G5:H5")
    Set MyButton = ActivitiesSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "TabulateChecked"
        .Caption = "Tabulate Practice"
    End With
    
    'Tabulate All
    Set ButtonRange = ActivitiesSheet.Range("G4:H4")
    Set MyButton = ActivitiesSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "TabulateAll"
        .Caption = "Tabulate All"
    End With

    'Clear saved button
    Set ButtonRange = ActivitiesSheet.Range("E2:F2")
    Set MyButton = ActivitiesSheet.Buttons.Add(ButtonRange.Left, ButtonRange.Top, ButtonRange.Width, ButtonRange.Height)
    With MyButton
        .OnAction = "ClearAllSaved"
        .Caption = "Clear Saved Practices"
    End With

    'Also put in the drop down menu
    Call ActivityDropDown(ActivitiesSheet, PracticeRange)
    
End Sub
Sub ActivityDropDown(NewSheet As Worksheet, DropDownRange As Range)
'Create a dropdown menu and autopoulate the indicated cell

    With DropDownRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ActivitiesList"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .InputMessage = ""
        .ErrorMessage = "Please choose from the drop-down list"
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub DateValidation(NewSheet As Worksheet, DateRange As Range)

    With DateRange.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="1/1/1990"
        .IgnoreBlank = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .ErrorMessage = "Please enter in a valid date"
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub CenterDropdown(NewSheet As Worksheet, CenterRange As Range)
'Make a dropdown list with center names in the indicated cell

    With CenterRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CenterNames"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .InputMessage = ""
        .ErrorMessage = "Please choose from the drop-down list"
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub ListLinks()

    Dim aLinks As Variant
    aLinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(aLinks) Then
        Sheets.Add
        For i = 1 To UBound(aLinks)
            Cells(i, 1).Value = aLinks(i)
        Next i
    End If
    
End Sub

Sub BreakExternalLinks()
'PURPOSE: Breaks all external links that would show up in Excel's "Edit Links" Dialog Box
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

Dim ExternalLinksArray As Variant
Dim wb As Workbook
Dim x As Long

Set wb = ActiveWorkbook

'Create an Array of all External Links stored in Workbook
  ExternalLinksArray = wb.LinkSources(Type:=xlLinkTypeExcelLinks)

'if the array is not empty the loop Through each External Link in ActiveWorkbook and Break it
 If IsEmpty(ExternalLinksArray) = False Then
     For x = 1 To UBound(ExternalLinksArray)
        wb.BreakLink Name:=ExternalLinksArray(x), Type:=xlLinkTypeExcelLinks
      Next x
End If

End Sub

