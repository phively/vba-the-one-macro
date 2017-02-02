Attribute VB_Name = "TheOneMacro_Dependent"
Option Explicit

Public C As CTimer ' These subs all get to share a CTimer because they will not interfere with the overall timings of the Core subs

Sub Timer(ByRef C As CTimer) ' Method to start the timer
    Set C = New CTimer
    C.StartCounter
End Sub

Sub RuntimeOptimization(ByVal bEvents As Boolean, ByVal bAlerts As Boolean, ByVal CalcMode As Long, ByVal bScreen As Boolean)

' Written by Paul Hively on 5/6/2013
' Turns off non-essential Excel functions during runtime to speed up macro performance
' Thanks to JP Software for the sample code: http://www.jpsoftwaretech.com/excel-vba/calculation-mode-excel-optimization/

' Disable Excel settings to speed up the macro
    With Application
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

End Sub

Sub RuntimeOptimizationOff(ByVal bEvents As Boolean, ByVal bAlerts As Boolean, ByVal CalcMode As Long, ByVal bScreen As Boolean)

' Written by Paul Hively on 5/6/2013; Last updated 6/3/2013
' Switches non-essential Excel functions back to their initial settings
' Thanks to JP Software for the sample code: http://www.jpsoftwaretech.com/excel-vba/calculation-mode-excel-optimization/

' Restore original Excel settings
    With Application
        .EnableEvents = bEvents
        .DisplayAlerts = bAlerts
        .Calculation = CalcMode
        .ScreenUpdating = bScreen
    End With
    
If DebugOn Then Debug.Print "Restoring Excel settings..."

End Sub

Sub LastUsedRow(ByRef rng As Range, row As Long)

' Written by Paul Hively on 2/4/2013
' Finds the last row number containing data in the range passed to this function

Dim col As Variant

' Reset row count
row = 0

' Loop through each column in the range
For Each col In rng.Columns
    ' If the current column has more rows than the previous record, this is the new maximum
    If col.Cells(rows.count, 1).End(xlUp).row > row Then
        row = col.Cells(rows.count, 1).End(xlUp).row
    End If
    ' Repeat until all columns are checked
Next col

End Sub

Sub LastUsedCol(ByRef rng As Range, col As Long)

' Written by Paul Hively on 2/4/2013
' Finds the last column number containing data in the row passed to this function

Dim row As Variant

' Reset column count
col = 0

' Loop through each row in the range
For Each row In rng.rows
    ' If the current row has more columns than the previous record, this is the new maximum
    If row.Cells(1, Columns.count).End(xlToLeft).Column > col Then
        col = row.Cells(1, Columns.count).End(xlToLeft).Column
    End If
    ' Repeat until all columns are checked
Next row

End Sub

Sub FormatInCell(ByRef rngData As Range, ArraySeek As Variant, Optional isBold As Boolean, Optional isItalic As Boolean, Optional isUnderlined As Boolean, Optional isColor As Boolean, Optional Color As Long)

' Written by Paul Hively on 10/1/2012; Last updated 6/19/2013
' Looks through the cell values of the specified range, and formats strings matching anything inside ArraySeek
   
' Format the matched names bold for easier checking
    Dim rngColumn As Range
    Dim rngCell As Range
    Dim Name As Variant
    Dim LookupName As Variant
    Dim intFoundText As Integer
    
' Iterate to the last value of this column
    For Each rngColumn In rngData.Columns
    ' Search through each cell in the range
        For Each rngCell In Range(rngColumn.Cells(1, 1), rngColumn.Cells(rows.count, 1).End(xlUp))
            ' Search through each name in the list
            For Each Name In ArraySeek
                ' Is the name string found in the curent cell?
                LookupName = CStr(Name)
                intFoundText = InStr(1, rngCell, LookupName, vbTextCompare)
                ' If found, format the name from its first appearance to its length
                If intFoundText > 0 Then
                    ' Format bold, italic, or underlined, if desired
                    If isBold Then rngCell.Characters(intFoundText, Len(LookupName)).Font.Bold = True
                    If isItalic Then rngCell.Characters(intFoundText, Len(LookupName)).Font.Italic = True
                    If isUnderlined Then rngCell.Characters(intFoundText, Len(LookupName)).Font.Underline = True
                    ' Format color
                    If isColor Then
                        If Color = 0 Then Color = RGB(255, 0, 0)
                        rngCell.Characters(intFoundText, Len(LookupName)).Font.Color = Color
                        rngCell.Characters(intFoundText, Len(LookupName)).Font.TintAndShade = 0
                    End If
                End If
            ' Check the next staff name
            Next
        ' Check the next cell
        Next
    ' Check the next column
    Next
    
' Move back to the originally selected cells to be extra-tidy
    
    rngData.Select
    
End Sub

Sub AutoFitWorksheet(Optional skipRow As Boolean, Optional skipCol As Boolean, Optional ByRef CWS As Worksheet)

' Written by Paul Hively on 10/1/2012
' Updated 12/8/2014 to make row and column resizing optional
' Fits all rows and columns on the current worksheet

    ' Auto Fit each row and column
    If skipCol <> True Then ActiveSheet.UsedRange.Columns.AutoFit
    If skipRow <> True Then ActiveSheet.UsedRange.rows.AutoFit
    ' Do it again in case we can squeeze out any extra space
    'ActiveSheet.UsedRange.Columns.AutoFit
    'ActiveSheet.UsedRange.rows.AutoFit
    ' Scroll up and put pointer back in the upper left to be tidy
    Application.GoTo ActiveSheet.Range("A1"), True

End Sub

Sub FontFormat(ByRef myRng As Range)

' Written by Paul Hively on 1/7/2013; Last updated 6/3/2013
' Formats the range passed to this sub to Calibri 9, centered horizontally and vertically, and wrapped.

' Profiling
If DebugOn Then
    Timer C:=C
    Debug.Print "    +Font Start: " & 0 & " ms"
End If

    ' Font settings
    With myRng.Font
        .Name = "Calibri"
        .size = 9

If DebugOn Then Debug.Print "     Font & Size: " & C.TimeElapsed & " ms"
    
    End With
    
    ' Alignment settings
    With myRng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

If DebugOn Then Debug.Print "     Font Align: " & C.TimeElapsed & " ms"
        
        .WrapText = True
                
If DebugOn Then Debug.Print "    -Font Wrap: " & C.TimeElapsed & " ms"
    
    End With

End Sub

Sub HeaderFormat(ByRef myRng As Range)

' Written by Paul Hively on 9/6/2012; last updated 1/7/2013
' Formats the range passed to this sub to have a maroon background with bold white text.
' myRange = range of cells to be formatted

    ' Bold and white font
    With myRng
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    ' Purple background
        .Interior.Color = RGB(78, 42, 132)
    End With
    ' Autofilter the Header range
    ActiveSheet.AutoFilterMode = False
    myRng.AutoFilter

End Sub

Sub HighlightCols(ByRef myArr As Variant, rows As Long)

' Written by Paul Hively on 9/13/2012; last updated 9/13/2012
' Takes the array passed to this sub and highlights columns with the indicated headers yellow
' myArr = array of header titles; rows = last row number

Dim i As Variant
Dim col As Long
Dim rng As Range

    ' Loop through the column names in myArr
    For Each i In myArr
        Set rng = Nothing
        With ActiveSheet
            ' Make the column yellow if it's found, otherwise go to next
            On Error Resume Next
            col = Cells.Find(i, , xlValues, xlWhole).Column
            Set rng = .Range(Cells(2, col).Address, Cells(rows, col).Address)
            On Error GoTo 0
            If Not rng Is Nothing Then rng.Interior.Color = 65535
        End With
    Next i

End Sub

Sub DeleteRows(ByRef myArr As Variant, myCol As String, rows As Long, cols As Long)

' Written by Paul Hively on 9/13/2012; last updated 9/13/2012
' Thanks to Silviu for his "Delete_with_Autofilter_Array" macro
' Takes the array passed to this sub and deletes rows that match criteria found in the indicated column
' myArr = array of criteria; myCol = name of column to filter;
'   rows = last row number to filter; cols = last column number to filter

Dim colNum As Integer
Dim i As Variant
Dim rng As Range
    
    With ActiveSheet
        ' Remove any existing AutoFilter
        .AutoFilterMode = False
        ' Find column name we will be working with in the active worksheet
        colNum = Cells.Find(myCol, , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Column
        ' Sequentially filter chosen column for each criteria in array
        For Each i In myArr
            .Range(Cells(1, 1).Address, Cells(rows, cols).Address).AutoFilter Field:=colNum, Criteria1:=i
            ' Delete any cells that were found matching the current Criteria
            Set rng = Nothing
            With .AutoFilter.Range
                ' Skips step if no match is found
                On Error Resume Next
                Set rng = .Offset(1, 0).Resize(.rows.count - 1, 1).SpecialCells(xlCellTypeVisible)
                ' Turn error checking back on
                On Error GoTo 0
                ' If matches then delete the current filtered rows
                If Not rng Is Nothing Then rng.EntireRow.Delete
            End With
        Next i
        ' Remove AutoFilter from worksheet
        .AutoFilterMode = False
    End With

End Sub

Sub DeleteCols(ByRef myArr As Variant)

' Written by Paul Hively on 9/13/2012; last updated 9/13/2012
' Takes the range passed to this sub and deletes columns with the indicated headers
' myArr = array of header titles

Dim colNum As Integer
Dim i As Variant
Dim rng As Range
    
    ' Loop through the column names in myArr
    For Each i In myArr
        ' Delete any column names that are found
        Set rng = Nothing
        With ActiveSheet
            ' Delete any columns found matching the current criteria
            On Error Resume Next
            Set rng = Cells.Find(i, , xlValues, xlWhole).EntireColumn
            ' Turn error checking back on
            On Error GoTo 0
            ' If matches then delete the current column
            If Not rng Is Nothing Then rng.EntireColumn.Delete
        End With
    Next i

End Sub

Sub PrintSettingsDep(Optional orient As String, Optional TopBot As Double, Optional LeftRight As Double, Optional HeadFoot As Double)

' Written by Paul Hively on 12/8/2014
' Breaking the Print Settings macro into a dependency so it can be called directly by other subs
' without going through the UserForm

' Profiling
If DebugOn Then
    Dim C As CTimer
    Set C = New CTimer
    C.StartCounter
    Debug.Print "*Print Start: " & 0 & " ms"
End If

' Default settings
    If orient = "" Then orient = xlLandscape 'default is landscape orientation, not portrait
    If TopBot = 0 Then TopBot = 0.75 'default is .75-inch margins
    If LeftRight = 0 Then LeftRight = 0.75 'default is .75-inch margins
    If HeadFoot = 0 Then HeadFoot = 0.4 'default is .4-inch header/footer

Application.PrintCommunication = False ' Turn off print communication to increase speed (factor of 30x)

With ActiveSheet.PageSetup
' Page settings
    .CenterHorizontally = True
    .CenterVertically = False
    .Orientation = orient 'variable
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = False

If DebugOn Then Debug.Print "  Page Settings: " & C.TimeElapsed & " ms"

' Margins - all variable
    .LeftMargin = Application.InchesToPoints(LeftRight)
    .RightMargin = Application.InchesToPoints(LeftRight)
    .TopMargin = Application.InchesToPoints(TopBot)
    .BottomMargin = Application.InchesToPoints(TopBot)
    .HeaderMargin = Application.InchesToPoints(HeadFoot)
    .FooterMargin = Application.InchesToPoints(HeadFoot)

If DebugOn Then Debug.Print "  Margins: " & C.TimeElapsed & " ms"

' Header/footer setup
    .ScaleWithDocHeaderFooter = True
    .LeftHeader = ""
    .LeftFooter = "Kellogg School of Management" & Chr(10) & "Confidential" 'Chr(10) is newline
Application.PrintCommunication = True 'Just in case it was disabled, since Header/Footers with &Variables can get truncated
Application.PrintCommunication = False
    .CenterHeader = Trim("&F" & Chr(10) & "&A") '&F is the filename; &A is the worksheet tab name
    .CenterFooter = Date & Chr(10) & "Prepared by " & Application.UserName 'Chr(10) is newline
Application.PrintCommunication = True 'Just in case it was disabled, since Header/Footers with &Variables can get truncated
Application.PrintCommunication = False
    .RightHeader = ""
    .RightFooter = Trim("Page " & "&P") '&P is the page number
End With

If DebugOn Then Debug.Print "  Header/Footer: " & C.TimeElapsed & " ms"

' Commit page setup changes
Application.PrintCommunication = True

If DebugOn Then
    Debug.Print "  Print Communication: " & C.TimeElapsed & " ms"
    Debug.Print "-Print End: " & C.TimeElapsed & " ms"
End If

End Sub

Sub SingleColumnToTab(ByRef CWS As Worksheet, rngData As Range, outputFormat As String, formatOutput As Boolean, parseSeparately As Boolean, parseBlanks As Boolean, overwriteWS As Boolean)

' Written by Paul Hively on 12/3/2012; Last updated 4/1/2013
' Thanks to http://www.ozgrid.com/VBA/item-worksheets.htm for the basic idea
' Filters the specified column and moves each data group to its own tab or workbook, with the option to skip blank cells.

' Variables
Dim rngCell As Range
Dim currentColName As String
Dim colNum As Integer
Dim colAddress As String
Dim currentItemName As String
Dim currentWSName As String
Dim rowCount As Long
Dim WSToKeep As Variant
Dim OriginalWorkbook As String

' Determine which worksheets and workbook name should not be overwritten
WSToKeep = Array("CurrUniqueList4Macro") ' Always keep the unique values list which will be created below
OriginalWorkbook = ActiveWorkbook.Name
' If the option to overwrite was unchecked
If overwriteWS = False Then
    ' Count the number of worksheets to be added and resize the array
    Dim WS As Worksheet
    ReDim Preserve WSToKeep(ActiveWorkbook.Worksheets.count + 1)
    ' Add all current worksheets to the array of names to keep
    Dim i As Long
    i = 2
    For Each WS In ActiveWorkbook.Worksheets
        WSToKeep(i) = WS.Name
        i = i + 1
    Next WS
End If

' Turn off any existing filters on the data sheet
    CWS.AutoFilterMode = False
        
' Initially set rngData to the entire column to be filtered
    currentColName = rngData.Range("A1").Value
    ' Find column name we will be working with in the active worksheet
    colNum = Cells.Find(currentColName, , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Column
    colAddress = CWS.Cells(1, colNum).Address
    
' Count rows in rngData
    rowCount = rngData.rows.count

' Recreate the "CurrUniqueList4Macro" worksheet for the current column
    On Error Resume Next
    Worksheets("CurrUniqueList4Macro").Delete
    Worksheets.Add().Name = "CurrUniqueList4Macro"
    On Error GoTo 0

'Filter the rngData range so only a unique list is created
    With Worksheets("CurrUniqueList4Macro")
        rngData.AdvancedFilter xlFilterCopy, , _
         Worksheets("CurrUniqueList4Macro").Range("A1"), True
        'Set a range variable to the unique list, less the heading.
        Set rngData = .Range("A2", .Range("A" & rowCount).End(xlUp))
    End With
        
' AutoFilter the original worksheet for each unique item in the dataset
    With CWS
        ' Display individual unique values on separate worksheets, if called for
        If parseSeparately = True Then
            ' Go through each unique value from the column
            For Each rngCell In rngData
            ' Skip any blanks if we are not using them
                If rngCell <> "" Then
                    ' Name of the worksheet to be created
                    currentItemName = rngCell
                    currentWSName = currentColName + "_" + currentItemName
                    ' Solve the >31 character worksheet name problem
                    If Len(currentWSName) > 31 Then
                        currentWSName = Left(currentColName, WorksheetFunction.Max(30 - Len(currentItemName), 15)) + "_" + Left(currentItemName, WorksheetFunction.Max(30 - Len(currentColName), 15))
                    End If
                ' Check that we're not overwriting a worksheet that needs to be kept
                    UniqueWSName WSName:=currentWSName, WSToKeep:=WSToKeep
                ' Do the filtering and copying
                    SingleColCopySub CWS:=CWS, colNum:=colNum, colAddress:=colAddress, currentItemName:=currentItemName, currentWSName:=currentWSName, outputFormat:=outputFormat, formatOutput:=formatOutput
                End If
            Next rngCell
        End If
        ' Display non-blank values together on one worksheet, if called for
        If parseSeparately = False Then
            ' Name of the worksheet to be created
            currentWSName = Left(currentColName, 26) + "_VALS"
        ' Check that we're not overwriting a worksheet that needs to be kept
            UniqueWSName WSName:=currentWSName, WSToKeep:=WSToKeep
        ' Do the filtering and copying
            SingleColCopySub CWS:=CWS, colNum:=colNum, colAddress:=colAddress, currentItemName:="<>", currentWSName:=currentWSName, outputFormat:=outputFormat, formatOutput:=formatOutput
        End If
        ' Display blank cells on separate worksheet, if called for
        If parseBlanks = True Then
            ' Name of the worksheet to be created
            currentWSName = Left(currentColName, 26) + "_BLNK"
        ' Check that we're not overwriting a worksheet that needs to be kept
            UniqueWSName WSName:=currentWSName, WSToKeep:=WSToKeep
        ' Do the filtering and copying
            SingleColCopySub CWS:=CWS, colNum:=colNum, colAddress:=colAddress, currentItemName:="", currentWSName:=currentWSName, outputFormat:=outputFormat, formatOutput:=formatOutput
        End If
    End With
            
' Clean up by deleting the temporary worksheet
        Windows(OriginalWorkbook).Activate
        Worksheets("CurrUniqueList4Macro").Delete

End Sub

Sub SingleColCopySub(ByRef CWS As Worksheet, colNum As Integer, colAddress As String, currentItemName As String, currentWSName As String, outputFormat As String, formatOutput As Boolean)

' Written by Paul Hively on 3/4/2013; last updated on 4/1/2013
' Sub for SingleColumToTab used to actually copy the range of interest to the new tab; split into its own sub for ease of maintenance

Dim WSH As Object
Dim TempPath As String

' Filter the chosen column in the original dataset for the current value
CWS.Range(colAddress).AutoFilter colNum, currentItemName

' If the output format is to separate tabs
If outputFormat = "T" Then
    ' Recreate a worksheet with the current item's name
    On Error Resume Next
        Worksheets(currentWSName).Delete
    On Error GoTo 0
    Worksheets.Add().Name = currentWSName
    ' Copy the visible filtered range(default of Copy Method) and leave behind hidden rows
    CWS.UsedRange.Copy Destination:=ActiveSheet.Range("A1")
    ' AutoFilter the created sheet
    Worksheets(currentWSName).Range("1:1").AutoFilter
End If

' If the output format is to separate files
If outputFormat = "F" Then
    ' Create path to the temp folder
    Set WSH = CreateObject("Scripting.FileSystemObject")
    TempPath = WSH.GetSpecialFolder(2) & "\"
    ' Create a workbook with the current item's name
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=TempPath & currentWSName
    ActiveWorkbook.ChangeFileAccess Mode:=xlReadOnly, WritePassword:="admin"
    ' Copy the visible filtered range
    CWS.UsedRange.Copy Destination:=ActiveSheet.Range("A1")
    'AutoFilter the created sheet
    ActiveSheet.Range("1:1").AutoFilter
End If

' If formatting with the WebI macro, run that sub now
If formatOutput Then
    WebiFormat
End If

End Sub


Sub UniqueWSName(ByRef WSName As String, WSToKeep As Variant)

' Written by Paul Hively on 3/4/2013; Last updated on 5/6/2013
' Checks to see whether the indicated string is already used as a worksheet name.
' If so, adds numbers to the end until a unique sheet name is found.

Dim WSKN As Variant ' To iterate through all WS names to be kept
Dim WSExists As Boolean
Dim suffix As Long ' Appended to the end of the proposed worksheet name
Dim invRep As String

suffix = 0

' Check for invalid characters and replace them with a valid one
invRep = "-"
WSName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(WSName, ":", invRep), "/", invRep), "\", invRep), "?", invRep), "*", invRep), "[", invRep), "]", invRep)

' Iterates through workbook to check if the worksheet name exists
WSExists:
    WSExists = False
    For Each WSKN In WSToKeep
    ' When a match is found, stop iterating
        If WSKN = WSName Then
            WSExists = True
            Exit For
        End If
    Next WSKN

' If the worksheet name does exist, try a new name
If WSExists Then
    suffix = suffix + 1
    WSName = Left(WSName, 31 - Len(CStr(suffix))) + CStr(suffix)
    GoTo WSExists:
End If

' Add the newly approved WSName to WSToKeep
ReDim Preserve WSToKeep(0 To UBound(WSToKeep) + 1)
WSToKeep(UBound(WSToKeep)) = WSName

End Sub



