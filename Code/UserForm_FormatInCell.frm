VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_FormatInCell 
   Caption         =   "Font Formatting Macro Options"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4320
   OleObjectBlob   =   "UserForm_FormatInCell.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_FormatInCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

' Clear the tag checking whether it's been run
    Me.Tag = ""

' Set default options for the macro
    ToggleButton_Bold.Value = True
    ToggleButton_Italic.Value = False
    ToggleButton_Underline.Value = False
    ' Default color to use
    Me.Button_ColorPicker.ForeColor = UserFormLong
    ' Selector should be set to whatever range is currently highlighted
    refedit_Selector.Value = Selection.Address

End Sub

Private Sub Button_ColorPicker_Click()

UserFormLong = PickNewColor(UserFormLong)

' Change the font color to whatever was just clicked
Me.Button_ColorPicker.ForeColor = UserFormLong

End Sub

Private Sub Button_RunMacro_Click()

' Toggle the RunMacro variable and hide the UserForm
    Me.Tag = "RanMacro"
    Me.Hide

End Sub
