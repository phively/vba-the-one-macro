VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_PrintSettings 
   Caption         =   "Print Settings Options"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   OleObjectBlob   =   "UserForm_PrintSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_PrintSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_RunMacro_Click()

' Toggle the RunMacro variable and hide the UserForm
    Me.Tag = "RanMacro"
    UserForm_PrintSettings.Hide

End Sub

Private Sub UserForm_Initialize()

' Clear the tag checking whether it's been run
    Me.Tag = ""

' Populate the list box; if renaming here, must also rename in the below RunMacro section
    combo_Presets.AddItem "Fit on one page"
    combo_Presets.AddItem "Fit on one page + text scaling"
    combo_Presets.AddItem "Custom settings"

' When the Print Settings Macro is run, use default settings
    combo_Presets.Value = "Fit on one page"
    combo_Presets_Change

End Sub

Private Sub madeChanges()
    combo_Presets.Value = "Custom settings"
End Sub

Private Sub combo_Presets_Change()

' If we're using custom settings then no action is required
If combo_Presets.Value = "Custom settings" Then Exit Sub
' If we're using a preset then we need to fill in the appropriate values here
' Fit on one page presets
If combo_Presets.Value = "Fit on one page" Then
    oLandscape.Value = True
    tTopBot.Value = 0.75
    tLeftRight.Value = 0.75
    tHeadFoot.Value = 0.4
    cFontRescaling.Value = False
    tFontRescaling.Value = ""
    cWordWrap.Value = True
    cFixedCols.Value = False
    o11x17.Value = True
    combo_Presets.Value = "Fit on one page"
End If
' Fit on one page + text scaling preset
If combo_Presets.Value = "Fit on one page + text scaling" Then
    oLandscape.Value = True
    tTopBot.Value = 0.1
    tLeftRight.Value = 0.15
    tHeadFoot.Value = 0.2
    cFontRescaling.Value = True
    tFontRescaling.Value = 9
    cWordWrap.Value = True
    cFixedCols.Value = True
    o11x17.Value = True
    combo_Presets.Value = "Fit on one page + text scaling"
End If

End Sub

Private Sub cWordWrap_Click()
    madeChanges
End Sub

Private Sub o11x17_Click()
    madeChanges
End Sub

Private Sub oLandscape_Click()
    madeChanges
End Sub

Private Sub oLegal_Click()
    madeChanges
End Sub

Private Sub oLetter_Click()
    madeChanges
End Sub

Private Sub oPortrait_Click()
    madeChanges
End Sub

Private Sub tFontRescaling_Change()
    cFontRescaling.Value = True
    If tFontRescaling.Value = "" Then cFontRescaling.Value = False
    madeChanges
End Sub

Private Sub tLeftRight_Change()
    madeChanges
End Sub

Private Sub tTopBot_Change()
    madeChanges
End Sub

Private Sub tHeadFoot_Change()
    madeChanges
End Sub

Private Sub cFontRescaling_Click()
    If cFontRescaling.Value = False Then tFontRescaling.Value = ""
    madeChanges
End Sub

Private Sub cFixedCols_Click()
    madeChanges
End Sub

