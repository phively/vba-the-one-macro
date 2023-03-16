Attribute VB_Name = "Booth_Core"
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
   FormatInCell rngData:=Selection, ArraySeek:=BoothAlumDevStaff, isBold:=True

If DebugOn Then
    Debug.Print "  Format In Cell: " & C.TimeElapsed & " ms"
    Debug.Print "-Find Phone End: " & C.TimeElapsed & " ms"
End If

' Switch Excel settings back to initial values
    If Reset Then RuntimeOptimizationOff bEvents:=bEvents, bAlerts:=bAlerts, CalcMode:=CalcMode, bScreen:=bScreen

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
        FormatInCell rngData:=Selection, ArraySeek:=myArray, isBold:=True, isColor:=True
    
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
Dim rngData As Range
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
    Set rngData = ActiveSheet.Range("N2", "N" & rowCount - 1)
    ' Iterate through each row after the header
    i = 2
    For Each rngCell In rngData
        ActiveSheet.Hyperlinks.Add Anchor:=Range("n" & i), Address:= _
        Range("l" & i).Value, TextToDisplay:=Range("j" & i).Value
        i = i + 1
    Next
    ' Special formatting (green) of key words within the hyperlink, if any
    Cells.Find("News", , xlValues, xlWhole).EntireColumn.Select
    FormatInCell rngData:=Selection, ArraySeek:=Array("trustee", "trustees", "board of "), isBold:=True, isColor:=True, Color:=RGB(0, 176, 80)
    
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
