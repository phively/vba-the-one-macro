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

Sub PrintSettings()

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
        ActiveSheet.PageSetup.PaperSize = xlPaper11x17
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

Sub MailerFormat()

' Written by Paul Hively on 9/4/2012; Last updated 6/3/2013
' Formats files from Mailer.

DebugOptions

' Variables
Dim rowCount As Long
Dim colCount As Long
Dim myArray As Variant
Dim myRange As Range

' Optimize Excel settings to speed up the macro
    Dim Reset As Boolean
    Reset = Application.ScreenUpdating
    SaveCurrentSettings bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen, Reset:=Reset
    RuntimeOptimization bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
    
' Profiling
If DebugOn Then
    Dim C As CTimer
    Set C = New CTimer
End If
    
' Initialize counts
LastUsedRow rng:=ActiveSheet.Range("A:A"), row:=rowCount
LastUsedCol rng:=ActiveSheet.Range("1:1"), col:=colCount

' Set Mailer Macro options (see UserForm_Mailer subs)
UserForm_Mailer.Show

' If Run Mailer Macro was clicked
    If UserForm_Mailer.Tag = "RanMacro" Then

If DebugOn Then
    C.StartCounter
    Debug.Print "*Mailer Start: " & 0 & " ms"
End If

    ' If Postal Mail type, then AutoFilter for blanks in LINE_1 and delete any that are found
        If UserForm_Mailer.OptionButton_Postal.Value = True Then
            myArray = Array("=")
            DeleteRows myArr:=myArray, myCol:="LINE_1", rows:=rowCount, cols:=colCount
            ' Refresh row count now that we've deleted some
            LastUsedRow rng:=ActiveSheet.Range("A:A"), row:=rowCount
If DebugOn Then Debug.Print "  Delete Blank Line 1: " & C.TimeElapsed & " ms"
        End If
    

    
    ' Delete Seeds, if the option was selected
    If UserForm_Mailer.OptionButton_DeleteSeedsY.Value = True Then
        ' Delete seed IDs; add or change IDs in the below variable
        ' List updated on 11/5/2012
            myArray = Array("5834458", "0005834458", "5911039", "0005911039", "6001120", "0006001120", "6257901", "0006257901", "6393527", "0006393527", "6525741", "0006525741", "6702133", "0006702133", "6787377", "0006787377", "6812464", "0006812464", "6923245", "0006923245", "6970679", "0006970679", "1000086356", "1000209673", "1000274968", "1000286759")
            DeleteRows myArr:=myArray, myCol:="ID_NUMBER", rows:=rowCount, cols:=colCount
            ' Refresh row count now that we've deleted some
            LastUsedRow rng:=ActiveSheet.Range("A:A"), row:=rowCount
If DebugOn Then Debug.Print "  Delete Seeds: " & C.TimeElapsed & " ms"
    End If
    
    ' Delete extra columns by name
        myArray = Array("ORG_CONTACT_NAME", "ORG_CONTACT_TITLE", "PHONE_TYPE", "PHONE", "MOBILE_PHONE", "EMAIL", "FAC_EX_BUILDING", "POSTNET_ZIP", "BARCODING_STREET", "RECORD_STATUS_CODE", "SPOUSE_REPORT_NAME", "FIRST_NAME", "MIDDLE_NAME", "LAST_NAME", "RIGHT_DATA")
        DeleteCols myArr:=myArray
        ' Refresh column count now that we've deleted some
        LastUsedCol rng:=ActiveSheet.Range("1:1"), col:=colCount
    
If DebugOn Then Debug.Print "  Delete Cols: " & C.TimeElapsed & " ms"
    
    ' Highlight the usable mailer fields in obnoxious yellow
        myArray = Array("SALUTATION", "LINE_1", "LINE_2", "LINE_3", "LINE_4", "LINE_5", "LINE_6", "LINE_7", "LINE_8")
        HighlightCols myArr:=myArray, rows:=rowCount

If DebugOn Then Debug.Print "  Highlight Cols: " & C.TimeElapsed & " ms"

    ' Look for salutations with the double apostrophe issue
        ActiveSheet.rows(1).Find("SALUTATION").Select
        myArray = Array("''")
        FormatInCell RngData:=Selection, ArraySeek:=myArray, isBold:=True, isColor:=True
    
If DebugOn Then Debug.Print "  Double Apostrophe: " & C.TimeElapsed & " ms"
    
    ' Set header shading
        Set myRange = ActiveSheet.Range(Cells(1, 1).Address, Cells(1, colCount).Address)
        HeaderFormat myRng:=myRange
        
If DebugOn Then Debug.Print "  Header Format: " & C.TimeElapsed & " ms"
        
    ' Run the WebI formatting macro
        WebiFormat

If DebugOn Then
    Debug.Print "   WebI Format: " & C.TimeElapsed & " ms"
    Debug.Print "-Mailer End: " & C.TimeElapsed & " ms"
End If

    End If
    
    Unload UserForm_Mailer
    
' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
 
End Sub

Sub NewsAlertsFormat()
'
' Written by Paul Hively on 10/1/2012; Last updated 6/19/2013
' Updated on 1/7/2013 to use new footer macro; future update should tidy up the header formatting
' Formats the Booth Alumni News Alerts output
' Be sure to check for valid URLs before running

DebugOptions

' Variables
Dim rowCount As Long
Dim RngData As Range
Dim rngCell As Range
Dim i As Long
Dim WS As Worksheet
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
    Debug.Print "*News Alerts Start: " & 0 & " ms"
End If
    
' Initialize count
LastUsedRow rng:=ActiveSheet.Range("A:A"), row:=rowCount

If DebugOn Then Debug.Print "  Row Counts: " & C.TimeElapsed & " ms"

' Format the staff names bold for easier checking
    Range(Cells.Find("URM", , xlValues, xlWhole).Address, Cells.Find("Team Managers", , xlValues, xlWhole).Address).Select
    FindPhoneList
 
If DebugOn Then Debug.Print "  Find Phone List: " & C.TimeElapsed & " ms"
 
' Puts the Hyperlink formula into each column in News
    Set RngData = ActiveSheet.Range("N2", "N" & rowCount - 1)
    ' Iterate through each row after the header
    i = 2
    For Each rngCell In RngData
        ActiveSheet.Hyperlinks.Add Anchor:=Range("n" & i), Address:= _
        Range("l" & i).Value, TextToDisplay:=Range("j" & i).Value
        i = i + 1
    Next
    ' Special formatting (green) of key words within the hyperlink, if any
    Cells.Find("News", , xlValues, xlWhole).EntireColumn.Select
    FormatInCell RngData:=Selection, ArraySeek:=Array("trustee", "trustees", "board of "), isBold:=True, isColor:=True, Color:=RGB(0, 176, 80)
    
If DebugOn Then Debug.Print "  Hyperlinks: " & C.TimeElapsed & " ms"
    
' Turn entire Record Types column to pretty maroon
    HeaderFormat myRng:=ActiveSheet.Range(Range("A1").End(xlToRight).Find("Record Types"), Range("A1").End(xlToRight).Find("Record Types").End(xlDown))
    
If DebugOn Then Debug.Print "  Column Highlight: " & C.TimeElapsed & " ms"
    
' Turn entire header row to pretty maroon
    HeaderFormat myRng:=ActiveSheet.Range("A1", Range("A1").End(xlToRight))
    ' Turn the columns to delete back to blue
    Range(Cells.Find("BOOL Foreign country?", , xlValues, xlWhole).Address, Cells.Find("BOOL Link isn't http", , xlValues, xlWhole).Address).Select
    Selection.Interior.Color = 16711680
   
If DebugOn Then Debug.Print "  Header Format: " & C.TimeElapsed & " ms"
   
' Filter, center, and set to Calibri 9
    FontFormat myRng:=ActiveSheet.UsedRange
    
' Make link columns font size 2
    Range(Cells.Find("Research Rpt N.B.", , xlValues, xlWhole).Address, Cells.Find("Research Rpt Linkable", , xlValues, xlWhole).Address).EntireColumn.Select
    Selection.Font.size = 2
    Range(Cells.Find("Research Rpt N.B.", , xlValues, xlWhole).Address, Cells.Find("Research Rpt Linkable", , xlValues, xlWhole).Address).Font.size = 9
    
If DebugOn Then Debug.Print "  Font Format: " & C.TimeElapsed & " ms"
    
' Set up header and footer
    PrintSettingsDep
    
If DebugOn Then Debug.Print "  Print Settings: " & C.TimeElapsed & " ms"

' Freeze the top row
    Range("2:2").Select
    With ActiveWindow
        .FreezePanes = False
        .FreezePanes = True
    End With

' Autofit and autofilter
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Range("A1").AutoFilter
    AutoFitWorksheet

If DebugOn Then
    Debug.Print "  Autofit Worksheet: " & C.TimeElapsed & " ms"
    Debug.Print "-News Alerts End: " & C.TimeElapsed & " ms"
End If

' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

End Sub

Sub FindPhoneList()

' Written by Paul Hively on 6/22/2012; Staff list last updated 1/21/2015
' Bolds all names that appear in the staff phone list BoothAlumDevStaff(1) to (N) in the selected columns
' Only need to select the first cell in each column to be bolded; this should be done in another sub.
   
DebugOptions

' Optimize Excel settings to speed up the macro
    Dim Reset As Boolean
    Reset = Application.ScreenUpdating
    SaveCurrentSettings bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen, Reset:=Reset
    RuntimeOptimization bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen
  
' Profiling
If DebugOn Then
    Dim C As CTimer
    Set C = New CTimer
    C.StartCounter
    Debug.Print "*Find Phone Start: " & 0 & " ms"
End If
 
' Staff last name dimension. The higher number on the next line MUST be >= the number of staff members
    Dim BoothAlumDevStaff(1 To 82) As String
        
' Write something to the entire array
    Dim i As Long
    
    For i = 1 To UBound(BoothAlumDevStaff)
        BoothAlumDevStaff(i) = "NOT_INITIALIZED"
    Next i
        
' Initialize staff last name list
' Added commas so we don't get partial matches, e.g. "Lee" matches "Kathleen"
        
BoothAlumDevStaff(1) = "Asaadi, H"
BoothAlumDevStaff(2) = "Axon, L"
BoothAlumDevStaff(3) = "Bada, L"
BoothAlumDevStaff(4) = "Bergmann, P"
BoothAlumDevStaff(5) = "Becka, J"
BoothAlumDevStaff(6) = "Buck, J"
BoothAlumDevStaff(7) = "Burgess, L"
BoothAlumDevStaff(8) = "Burns, S"
BoothAlumDevStaff(9) = "Cernosia, P"
BoothAlumDevStaff(10) = "Chamberlin, S"
BoothAlumDevStaff(11) = "Chan, K"
BoothAlumDevStaff(12) = "Coogan, K"
BoothAlumDevStaff(13) = "De bie, T"
BoothAlumDevStaff(14) = "De may, E"
BoothAlumDevStaff(15) = "Dentamaro, C"
BoothAlumDevStaff(16) = "Dilley, D"
BoothAlumDevStaff(17) = "Douponce, M"
BoothAlumDevStaff(18) = "Dove, C"
BoothAlumDevStaff(19) = "Eldringhoff, C"
BoothAlumDevStaff(20) = "Eriksson, J"
BoothAlumDevStaff(21) = "Eunson, L"
BoothAlumDevStaff(22) = "Foster, D"
BoothAlumDevStaff(23) = "Furlong, L"
BoothAlumDevStaff(24) = "Gonnella, L"
BoothAlumDevStaff(25) = "Green, A"
BoothAlumDevStaff(26) = "Griffin, T"
BoothAlumDevStaff(27) = "Griffith, L"
BoothAlumDevStaff(28) = "Guynn, L"
BoothAlumDevStaff(29) = "Haley, R"
BoothAlumDevStaff(30) = "Hassett, L"
BoothAlumDevStaff(31) = "Hill, P"
BoothAlumDevStaff(32) = "Hively, P"
BoothAlumDevStaff(33) = "Hoffmann, C"
BoothAlumDevStaff(34) = "Hollendoner, E"
BoothAlumDevStaff(35) = "Huml, M"
BoothAlumDevStaff(36) = "Humphrey, M"
BoothAlumDevStaff(37) = "Johnston, N"
BoothAlumDevStaff(38) = "Karr, C"
BoothAlumDevStaff(39) = "Kondrat, B"
BoothAlumDevStaff(40) = "Lalonde, A"
BoothAlumDevStaff(41) = "Lam, W"
BoothAlumDevStaff(42) = "Le, C"
BoothAlumDevStaff(43) = "Lee, K"
BoothAlumDevStaff(44) = "Macdougall, M"
BoothAlumDevStaff(45) = "Mahgoub, S"
BoothAlumDevStaff(46) = "Mccabe, T"
BoothAlumDevStaff(47) = "Mckee, R"
BoothAlumDevStaff(48) = "Mckenzie, M"
BoothAlumDevStaff(49) = "Mendoza, A"
BoothAlumDevStaff(50) = "Mihalek, J"
BoothAlumDevStaff(51) = "Miller, E"
BoothAlumDevStaff(52) = "Miller, L"
BoothAlumDevStaff(53) = "Mutchler, L"
BoothAlumDevStaff(54) = "Nash, R"
BoothAlumDevStaff(55) = "Nelson, A"
BoothAlumDevStaff(56) = "Niermann, A"
BoothAlumDevStaff(57) = "O'connor, K"
BoothAlumDevStaff(58) = "Olivier, T"
BoothAlumDevStaff(59) = "Porta, S"
BoothAlumDevStaff(60) = "Primlani, A"
BoothAlumDevStaff(61) = "Reid, C"
BoothAlumDevStaff(62) = "Ritchell, C"
BoothAlumDevStaff(63) = "Rodriguez, J"
BoothAlumDevStaff(64) = "Saleh, M"
BoothAlumDevStaff(65) = "Scheid, V"
BoothAlumDevStaff(66) = "Schneider, K"
BoothAlumDevStaff(67) = "Schreur, E"
BoothAlumDevStaff(68) = "Seeley, P"
BoothAlumDevStaff(69) = "Seyal, T"
BoothAlumDevStaff(70) = "Shafaee, M"
BoothAlumDevStaff(71) = "Smith, A"
BoothAlumDevStaff(72) = "Smith, K"
BoothAlumDevStaff(73) = "State, T"
BoothAlumDevStaff(74) = "Su, C"
BoothAlumDevStaff(75) = "Sullivan, K"
BoothAlumDevStaff(76) = "Tiemens, C"
BoothAlumDevStaff(77) = "Tkach, K"
BoothAlumDevStaff(78) = "Trent, A"
BoothAlumDevStaff(79) = "Ute, L"
BoothAlumDevStaff(80) = "Westhouse, E"
BoothAlumDevStaff(81) = "Yoo, S"
BoothAlumDevStaff(82) = "Young, K"

If DebugOn Then Debug.Print "  Array Created: " & C.TimeElapsed & " ms"

' Call the FormatInCell sub
   FormatInCell RngData:=Selection, ArraySeek:=BoothAlumDevStaff, isBold:=True

If DebugOn Then
    Debug.Print "  Format In Cell: " & C.TimeElapsed & " ms"
    Debug.Print "-Find Phone End: " & C.TimeElapsed & " ms"
End If

' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

End Sub

Sub ColumnsToTabs()

' Written by Paul Hively on 12/3/2012; Last updated 6/3/2013
' Iterate through a set of columns to be filtered with the SingleColumnToTab sub

DebugOptions

' Variables
Dim CWS As Worksheet
Dim RngData As Range
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
        Set RngData = Range(UserForm_ColToTab.refedit_Selector.Value)
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
        For Each rngColumn In RngData.Columns
            ' Run the SingleColumnToTab dependent sub for the current column
            SingleColumnToTab CWS:=CWS, RngData:=rngColumn, outputFormat:=UserForm_ColToTab.combo_OutputFormat, formatOutput:=UserForm_ColToTab.chkbox_RunWebIFormat, parseSeparately:=UserForm_ColToTab.chkbox_parseSeparately.Value, parseBlanks:=UserForm_ColToTab.chkbox_parseBlanks.Value, overwriteWS:=UserForm_ColToTab.chkbox_OverwriteWS
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
Dim RngData As Range
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
            Set RngData = Range(UserForm_FormatInCell.refedit_Selector.Value)
        On Error GoTo 0
    
    If DebugOn Then Debug.Print " RngData: " & RngData.Address
    
    ' Run Format In Cell
        FormatInCell RngData:=RngData, ArraySeek:=Array(UserForm_FormatInCell.TextBox_SearchString.Value), isBold:=UserForm_FormatInCell.ToggleButton_Bold.Value, isItalic:=UserForm_FormatInCell.ToggleButton_Italic.Value, isUnderlined:=UserForm_FormatInCell.ToggleButton_Underline, isColor:=True, Color:=UserFormLong
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

Sub GreenSheetFormat()

' Started by Paul Hively on 1/31/2017
' Automatic ColumnsToTabs for gift officer "Green Sheet" reports

' ************ OK TO EDIT BELOW HERE ************
' Prospect managers dimension. The higher number on the next line MUST be >= the number of staff members
    Dim KSMProspectManagers(1 To 20) As String
' Write something to the entire array
    Dim i As Long
    For i = 1 To UBound(KSMProspectManagers)
        KSMProspectManagers(i) = "NOT_INITIALIZED"
    Next i
' Insert the prospect managers into the array
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

' ************ DO NOT EDIT BELOW HERE ************

End Sub

