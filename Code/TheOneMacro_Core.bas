Attribute VB_Name = "TheOneMacro_Core"
Option Explicit

Global DebugOn As Boolean

' Used to save current Excel settings before changing to optimized settings
Public bEvents As Boolean
Public bAlerts As Boolean
Public CalcMode As Long
Public bScreen As Boolean

' Variables to be used in UserForms
Public UserFormLong As Long

Private Sub Workbook_Open()
' Written by Paul Hively on 5/4/2022: set calculation mode to xlAutomatic when opening Excel
Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub DebugOptions()
    DebugOn = True ' If set to true, debug info will display in the Immediate window
End Sub

Public Sub SaveCurrentSettings(ByRef bEvents As Boolean, ByRef bAlerts As Boolean, ByRef CalcMode As Long, ByRef bScreen As Boolean, ByRef Reset As Boolean)
' Written by Paul Hively on 6/3/2013; split out from the individual subs below
' If Excel settings have not been stored this session, write them now

If Not Reset Then Exit Sub

If DebugOn Then Debug.Print "Writing Excel settings..."
    bEvents = Application.EnableEvents
    bAlerts = Application.DisplayAlerts
    CalcMode = Application.Calculation
    bScreen = Application.ScreenUpdating
    
End Sub

Sub WebiFormat()

' Written by Paul Hively on 1/7/2013; last updated 10/7/2013
' Refactored from original WebiFormat recorded macro.
' Formats WebI output files to match the Biodata & Analytics format. User Name is set in Options -> Personalize.

DebugOptions

' Optimize Excel settings to speed up the macro
    Dim Reset As Boolean
    Reset = Application.ScreenUpdating
    SaveCurrentSettings bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen, Reset:=Reset
    RuntimeOptimization bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

' Variables
    Dim rowCount As Long
    Dim colCount As Long
    Dim myRange As Range

' Profiling
If DebugOn Then
    Dim C As CTimer
    Set C = New CTimer
    C.StartCounter
    Debug.Print "*WebI Start: " & 0 & " ms"
End If

' Initialize counts
    LastUsedRow rng:=ActiveSheet.Range("A:A"), row:=rowCount
    LastUsedCol rng:=ActiveSheet.Range("1:1"), col:=colCount
    
If DebugOn Then Debug.Print "  Row Col Counts: " & C.TimeElapsed & " ms"
    
' Header shading and AutoFilter
    Set myRange = ActiveSheet.Range(Cells(1, 1).Address, Cells(1, colCount).Address)
    HeaderFormat myRng:=myRange

If DebugOn Then Debug.Print "  Header Shading: " & C.TimeElapsed & " ms"

' Set font, font size, alignment, and wrapping
    FontFormat myRng:=ActiveSheet.UsedRange

If DebugOn Then Debug.Print "  Font Format: " & C.TimeElapsed & " ms"

' Page setup with footer and print settings
    PrintSettingsDep

If DebugOn Then Debug.Print "  Print Settings: " & C.TimeElapsed & " ms"

' Freeze the top row and print headers on every page
    ActiveSheet.PageSetup.PrintTitleRows = "$1:$1"
    ' Move the selection to A1 so we freeze the correct area
    Application.GoTo ActiveSheet.Range("A1"), True
    ' Everything above the selection will be frozen
    Range("2:2").Select
    With ActiveWindow
        .FreezePanes = False
        .FreezePanes = True
    End With
      
If DebugOn Then Debug.Print "  Freeze Panes: " & C.TimeElapsed & " ms"
      
' AutoFit the current worksheet
     ' We don't want the columns to be too wide so we set them to the desired width (with tall rows) beforehand for word wrap
    ActiveSheet.UsedRange.Columns.ColumnWidth = 60
    ActiveSheet.UsedRange.rows.RowHeight = 408
    AutoFitWorksheet
      
If DebugOn Then
    Debug.Print "  Autofit Worksheet: " & C.TimeElapsed & " ms"
    Debug.Print "-WebI End: " & C.TimeElapsed & " ms"
End If

' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
  
End Sub

Sub printSettings()

' Written by Paul Hively on 1/7/2013; updated 6/3/2013
' Sets the default print options used by Biodata & Analytics
' See http://support.microsoft.com/kb/142136 for full list of header/footer commands
' 12/9/2014 extensive rewrite - broke shared components into a dependent sub and added a UserForm
' 3/2/2015 added A4 paper size (thanks Penka)

DebugOptions

' Optimize Excel settings to speed up the macro
    Dim Reset As Boolean
    Reset = Application.ScreenUpdating
    SaveCurrentSettings bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen, Reset:=Reset
    RuntimeOptimization bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
    
' Set Mailer Macro options (see UserForm_Mailer subs)
UserForm_PrintSettings.Show

' If Run Mailer Macro was clicked
If UserForm_PrintSettings.Tag = "RanMacro" Then
       
    ' A few variables to be used later
    Dim orient As String 'page orientation
    Dim colCount As Long 'number of columns on the worksheet
    Dim width As Double 'to store width of the page in pixels
    Dim a As Double 'half-baked regression coefficients
    Dim b As Double 'half-baked regression coefficients
    Dim size As Double 'to store font size
    
    ' Force correct page orientation
    If UserForm_PrintSettings.oLandscape = True Then orient = xlLandscape
    If UserForm_PrintSettings.oPortrait = True Then orient = xlPortrait
    
    ' Run the basic page setup macro with the selected form options
    PrintSettingsDep orient, TopBot:=UserForm_PrintSettings.tTopBot, LeftRight:=UserForm_PrintSettings.tLeftRight, HeadFoot:=UserForm_PrintSettings.tHeadFoot

    ' Paper size selection
    If UserForm_PrintSettings.oLetter = True Then
        ActiveSheet.PageSetup.PaperSize = xlPaperLetter
        If orient = xlPortrait Then width = 8.5 - 2 * UserForm_PrintSettings.tLeftRight
        If orient = xlLandscape Then width = 11 - 2 * UserForm_PrintSettings.tLeftRight
    End If
    If UserForm_PrintSettings.oLegal = True Then
        ActiveSheet.PageSetup.PaperSize = xlPaperLegal
        If orient = xlPortrait Then width = 8.5 - 2 * UserForm_PrintSettings.tLeftRight
        If orient = xlLandscape Then width = 14 - 2 * UserForm_PrintSettings.tLeftRight
    End If
    If UserForm_PrintSettings.o11x17 = True Then
        ActiveSheet.PageSetup.PaperSize = xlPaperTabloid
        If orient = xlPortrait Then width = 11 - 2 * UserForm_PrintSettings.tLeftRight
        If orient = xlLandscape Then width = 17 - 2 * UserForm_PrintSettings.tLeftRight
    End If
    If UserForm_PrintSettings.oA4 = True Then
        ActiveSheet.PageSetup.PaperSize = xlPaperA4
        If orient = xlPortrait Then width = 8.27 - 2 * UserForm_PrintSettings.tLeftRight
        If orient = xlLandscape Then width = 11.7 - 2 * UserForm_PrintSettings.tLeftRight
    End If
    
    ' Word Wrap
    If UserForm_PrintSettings.cWordWrap = True Then ActiveSheet.UsedRange.WrapText = True
    
    ' Column and font rescaling
    If UserForm_PrintSettings.cFontRescaling = True Then
        
        ' Count number of columns in the spreadsheet
        LastUsedCol rng:=ActiveSheet.Range("1:12"), col:=colCount 'check first 12 rows to avoid issues with merged cells
        
        ' Figure out how wide columns can be to fill the page
        width = width / colCount 'width of each column if identically sized in inches
        
        ' Figure out optimal font size given column width
        ' Half-baked regression coefficients
        a = 15.649
        b = 7.171
        size = Round(a * width + b, 0)
        ActiveSheet.UsedRange.Select
        ' Use optimal or minimum size, whichever is larger
        If size >= UserForm_PrintSettings.tFontRescaling Then Selection.Font.size = size
        If size < UserForm_PrintSettings.tFontRescaling Then Selection.Font.size = UserForm_PrintSettings.tFontRescaling
        
        ' Resize columns to take advantage of the new font size
        ActiveSheet.UsedRange.ColumnWidth = 30 * width 'characters scaled based on width in inches
        
    End If

    ' Column and row fitting
    AutoFitWorksheet skipCol:=UserForm_PrintSettings.cFixedCols 'according to user preferences

End If

Unload UserForm_PrintSettings

' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

Exit Sub ' Leave before getting to the error handling entries

' Error handling - not active as of 12/9/2014
ErrNonNum:
    MsgBox "Non-numeric textbox entries"
    ' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
    Exit Sub

End Sub

Sub ColumnsToTabs()

' Written by Paul Hively on 12/3/2012; Last updated 6/3/2013
' Iterate through a set of columns to be filtered with the SingleColumnToTab sub

DebugOptions

' Variables
Dim CWS As Worksheet
Dim rngData As Range
Dim rngColumn As Range
Dim Reset As Boolean

' Optimize Excel settings to speed up the macro
    Reset = Application.ScreenUpdating
    SaveCurrentSettings bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen, Reset:=Reset
    RuntimeOptimization bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

' Set ColumnsToTabs options (See UserForm_ColToTab)
' Prompt to select columns of interest for filtering, as well as whether blanks should be ignored (default)
    UserForm_ColToTab.Show
    
' If Run Macro button was clicked
    If UserForm_ColToTab.Tag = "RanMacro" Then
    
    ' Remember the currently selected worksheet
        Set CWS = ActiveSheet
    
    ' Define the range to be used
    On Error GoTo ErrInvRange
        Set rngData = Range(UserForm_ColToTab.refedit_Selector.Value)
    On Error GoTo 0
    ' Check the output format
    If UserForm_ColToTab.combo_OutputFormat = "Err" Then GoTo ErrNoOutput
    
' Profiling
If DebugOn Then
    Dim C As CTimer
    Set C = New CTimer
    C.StartCounter
    Debug.Print "*Columns to Tabs Start: " & 0 & " ms"
End If
    
    ' For each specified column
        For Each rngColumn In rngData.Columns
            ' Run the SingleColumnToTab dependent sub for the current column
            SingleColumnToTab CWS:=CWS, rngData:=rngColumn, OutputFormat:=UserForm_ColToTab.combo_OutputFormat, formatOutput:=UserForm_ColToTab.chkbox_RunWebIFormat, parseSeparately:=UserForm_ColToTab.chkbox_parseSeparately.Value, parseBlanks:=UserForm_ColToTab.chkbox_parseBlanks.Value, overwriteWS:=UserForm_ColToTab.chkbox_OverwriteWS
                'Make the booleans pull from the user form; should default to true and true
        Next rngColumn

If DebugOn Then Debug.Print "  All Cols Done: " & C.TimeElapsed & " ms"

    ' Re-select original worksheet
        With CWS
            .AutoFilterMode = False
            .Activate
        End With

If DebugOn Then Debug.Print "-Columns to Tabs End: " & C.TimeElapsed & " ms"

    End If
        
    Unload UserForm_ColToTab

' Switch Excel settings back to initial values
    RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

Exit Sub

' Error handling
ErrInvRange:
    MsgBox "Invalid range"
    ' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
    Exit Sub
ErrNoOutput:
    MsgBox "Invalid output format"
    ' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
    Exit Sub

End Sub

Sub AutoFitWorkbook()

' Written by Paul Hively on 6/22/2012; last update 6/3/2013
' Fits all rows and columns on each worksheet in the entire workbook

DebugOptions

' Variables
Dim WS As Worksheet
Dim CWS As Integer
Dim Reset As Boolean

' Optimize Excel settings to speed up the macro
    Reset = Application.ScreenUpdating
    SaveCurrentSettings bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen, Reset:=Reset
    RuntimeOptimization bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

' Profiling
If DebugOn Then
    Dim C As CTimer
    Set C = New CTimer
    C.StartCounter
    Debug.Print "*Autofit Workbook Start: " & 0 & " ms"
End If

' Remember the currently selected worksheet
    CWS = ActiveSheet.Index

' Loop through each worksheet
    For Each WS In ActiveWorkbook.Worksheets
        ' Select the new worksheet and autofit it
        WS.Activate
        AutoFitWorksheet
    ' Go to the next worksheet in the workbook
    Next WS

If DebugOn Then Debug.Print "  All Worksheets Fit: " & C.TimeElapsed & " ms"

' Re-select the originally selected worksheet in the workbook
    Worksheets(CWS).Select

If DebugOn Then Debug.Print "-Autofit Workbook End: " & C.TimeElapsed & " ms"

' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

End Sub

Sub FormatTextInCell()

' Written by Paul Hively on 7/8/2013
' Front end for the FormatInCell sub; finds specified string and bolds, changes the color, etc.

DebugOptions

' Variables
Dim rngData As Range
UserFormLong = RGB(255, 0, 0) ' Default font color to use

' Optimize Excel settings to speed up the macro
    Dim Reset As Boolean
    Reset = Application.ScreenUpdating
    SaveCurrentSettings bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen, Reset:=Reset
    RuntimeOptimization bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

' Set FormatInCell options (See UserForm_FormatInCell)
    UserForm_FormatInCell.Show

' If Run Macro button was clicked
    If UserForm_FormatInCell.Tag = "RanMacro" Then
    
    ' Define the range to be used
        On Error GoTo ErrInvRange
            Set rngData = Range(UserForm_FormatInCell.refedit_Selector.Value)
        On Error GoTo 0
    
    If DebugOn Then Debug.Print " RngData: " & rngData.Address
    
    ' Run Format In Cell
        FormatInCell rngData:=rngData, ArraySeek:=Array(UserForm_FormatInCell.TextBox_SearchString.Value), isBold:=UserForm_FormatInCell.ToggleButton_Bold.Value, isItalic:=UserForm_FormatInCell.ToggleButton_Italic.Value, isUnderlined:=UserForm_FormatInCell.ToggleButton_Underline, isColor:=True, Color:=UserFormLong
    End If

    Unload UserForm_FormatInCell

' Switch Excel settings back to initial values
    RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

Exit Sub

' Error handling
ErrInvRange:
    MsgBox "Invalid range"
    ' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
    Exit Sub

End Sub


