Attribute VB_Name = "KSM_Core"
' Kellogg-specific macro code

Sub GreenSheetFormat()

' Started by Paul Hively on 1/31/2017
' Automatic ColumnsToTabs for gift officer "Green Sheet" reports

' *****************************************
' *********** EDIT BELOW HERE *************
' *****************************************
' Name of the prospect managers column
    Dim PMColName As String
' Prospect managers dimension. The higher number on the next line MUST be >= the number of staff members
    Dim KSMProspectManagers(1 To 20) As String
' Fields dimension; names of the columns that will be kept in the formatted spreadsheet
    Dim FieldsToKeep(1 To 20) As String
' Write something to the entire array
    Dim i As Long
    For i = 1 To UBound(KSMProspectManagers)
        KSMProspectManagers(i) = "NOT_INITIALIZED"
    Next i
    Dim j As Long
    For j = 1 To UBound(FieldsToKeep)
        FieldsToKeep(j) = "NOT_INITIALIZED"
    Next j

' ************ PROSPECT MANAGER COLUMN NAME
    PMColName = "PM"
' ************ PROSPECT MANAGERS -- ADD NAMES BELOW IN SAME FORMAT
    KSMProspectManagers(1) = "Ms. Erin Varga"
    KSMProspectManagers(2) = "Ms. Lisa Guynn"
    KSMProspectManagers(3) = "Ms. Sally Spritz"
    KSMProspectManagers(4) = "Mrs. Catherine C. Taylor"
    KSMProspectManagers(5) = "Mr. Adam Kristopher Nordmark"
    KSMProspectManagers(6) = "Ms. Suzanne K. Schoeneweiss"
    KSMProspectManagers(7) = "Mr. David S. Decker-Drane"
    KSMProspectManagers(8) = "Mr. Ryan Heath Jones"
    KSMProspectManagers(9) = "Ms. Janice Paszczykowski"
    KSMProspectManagers(10) = "Ms. Maggie T. Cong-Huyen"
    KSMProspectManagers(11) = "Mr. Jason Scott Keene"
    KSMProspectManagers(12) = "Ms. Jane Erb"
    KSMProspectManagers(13) = "Ms. Christine Kuhn Feary"
' ************ FIELDS TO KEEP -- ADD COLUMN NAMES BELOW IN SAME FORMAT
    FieldsToKeep(1) = "ID"
    FieldsToKeep(2) = "PREF_MAIL_NAME"
    FieldsToKeep(3) = "KSM YEAR"
    FieldsToKeep(4) = "Pref_KSM_Section"
    FieldsToKeep(5) = "KSM_Reunion_Group"
    FieldsToKeep(6) = "Ask_Amount"
    FieldsToKeep(7) = "City"
    FieldsToKeep(8) = "State"
    FieldsToKeep(9) = "Country"
    FieldsToKeep(10) = "EMPLOYER"
    FieldsToKeep(11) = "TITLE"
    FieldsToKeep(12) = "Kellogg Annual Giving"
    FieldsToKeep(13) = "Kellogg Annual Giving Year"
    FieldsToKeep(14) = "LIFETIME_GIVING_TOTAL"
    FieldsToKeep(15) = "PM"
' *****************************************
' ******** DO NOT EDIT BELOW HERE *********
' *****************************************

''' Macro and Excel settings
' Debug printing, if set in TheOneMacro_Core
    DebugOptions
' Profiling
    If DebugOn Then
        Dim C As CTimer
        Set C = New CTimer
        C.StartCounter
        Debug.Print "*GreenSheet Start: " & 0 & " ms"
    End If
' Optimize Excel settings to speed up the macro
    Dim Reset As Boolean
    Reset = Application.ScreenUpdating
    SaveCurrentSettings bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen, Reset:=Reset
' Comment out the next line when debugging to see the operations as they're carried out
    RuntimeOptimization bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen


''' Variables
Dim rowCount As Long
Dim colCount As Long
' For column/row deletion
Dim allCols As Variant
Dim allPMs As Variant
Dim thisOne As Variant
Dim colNum As Integer
Dim PMsCol As Range

' Initialize counts
LastUsedRow rng:=ActiveSheet.Range("A:A"), row:=rowCount
LastUsedCol rng:=ActiveSheet.Range("1:1"), col:=colCount

' Copy data to a new workbook
ActiveSheet.Copy Before:=ActiveSheet
ActiveSheet.Name = "WorkingSheet"


''' Delete unneeded columns
' First, store the header names
allCols = Range(Cells(1, 1), Cells(1, colCount)).Value
' Iterate through each named column; if not in FieldsToKeep then delete
For Each thisOne In allCols
    'If DebugOn Then Debug.Print thisOne
    If Not InArray(thisOne, FieldsToKeep) Then DeleteCols myArr:=Array(thisOne)
Next thisOne

' Profiling
    If DebugOn Then Debug.Print "  Delete columns: " & C.TimeElapsed & " ms"


''' Delete the PMs we don't need to see
' Create a worksheet to store the unique values
Sheets.Add.Name = "CurrUniqueList4Macro"
' Then, find and select the PM column
Worksheets("WorkingSheet").Activate
colNum = Cells.Find(PMColName, , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Column
Set PMsCol = ActiveSheet.Range(Cells(1, colNum), Cells(rowCount, colNum))
' Next, de-dupe the column on the new worksheet
    With Worksheets("CurrUniqueList4Macro")
        PMsCol.AdvancedFilter xlFilterCopy, , _
         Worksheets("CurrUniqueList4Macro").Range("A1"), True
        'Set a range variable to the unique list, less the heading.
        Set rngData = .Range("A2", .Range("A" & rowCount).End(xlUp))
    End With
' Finally, store these individuals in allPMs
allPMs = Worksheets("CurrUniqueList4Macro").UsedRange.Value
' Sort by PM Column name; this GREATLY reduces the amount of time it takes to delete rows
ActiveSheet.UsedRange.Sort key1:=PMsCol, Header:=xlYes
' Last, iterate through each PM name; if not in KSMProspectManagers then delete
For Each thisOne In allPMs
    'If DebugOn Then Debug.Print thisOne
    If Not InArray(thisOne, KSMProspectManagers) Then DeleteRows myArr:=Array(thisOne), myCol:=PMColName, rows:=rowCount, cols:=colCount
Next thisOne
' Clean up by deleting CurrUniqueList4Macro
Worksheets("CurrUniqueList4Macro").Delete

' Profiling
    If DebugOn Then Debug.Print "  Delete rows: " & C.TimeElapsed & " ms"


''' Split the file up into tabs
SingleColumnToTab CWS:=Worksheets("WorkingSheet"), rngData:=PMsCol, outputFormat:="T", _
    formatOutput:=True, parseSeparately:=True, parseBlanks:=False, overwriteWS:=True
' Clean up by deleting the temporary WorkingSheet tab
Worksheets("WorkingSheet").Delete

' Profiling
    If DebugOn Then Debug.Print "  Create tabs: " & C.TimeElapsed & " ms"


''' Cleanup
' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

' Profiling
    If DebugOn Then Debug.Print "*GreenSheet End: " & C.TimeElapsed & " ms"

End Sub
