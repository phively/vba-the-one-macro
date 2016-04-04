VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Mailer 
   Caption         =   "Mailer Macro Options"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   OleObjectBlob   =   "UserForm_Mailer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Mailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

' Clear the tag checking whether it's been run
    Me.Tag = ""

' When the Mailer Macro is run, default to postal mailing and keep seeds
    OptionButton_Postal.Value = True
    OptionButton_DeleteSeedsN.Value = True

End Sub

Private Sub Button_RunMacro_Click()

' Toggle the RunMacro variable and hide the UserForm
    Me.Tag = "RanMacro"
    UserForm_Mailer.Hide

End Sub
