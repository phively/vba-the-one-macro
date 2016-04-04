VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ColToTab 
   Caption         =   "Columns to Tabs Macro Options"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   OleObjectBlob   =   "UserForm_ColToTab.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_ColToTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

' Clear the tag checking whether it's been run
    Me.Tag = ""

' Populate the list box; if renaming here, must also rename in the below RunMacro section
    combo_OutputFormat.AddItem "Separate Tabs"
    combo_OutputFormat.AddItem "Separate Files"

' When the Columns to Tabs Macro is run, set default options
    combo_OutputFormat.Value = "Separate Tabs"
    chkbox_RunWebIFormat.Value = False
    chkbox_parseSeparately.Value = True
    chkbox_parseBlanks.Value = False
    chkbox_OverwriteWS = True
    ' Selector should be set to whatever range is currently highlighted
    refedit_Selector.Value = Selection.Address

End Sub

Private Sub RunMacro_Click()

' Convert the output format from a text field to a code
    If combo_OutputFormat.Value = "Separate Tabs" Then
        combo_OutputFormat.Value = "T"
    ElseIf combo_OutputFormat.Value = "Separate Files" Then
        combo_OutputFormat.Value = "F"
    Else: combo_OutputFormat.Value = "Err"
    End If

' Toggle the RunMacro variable and hide the UserForm
    Me.Tag = "RanMacro"
    UserForm_ColToTab.Hide

End Sub
