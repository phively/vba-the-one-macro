Attribute VB_Name = "KSM_Core"
' Kellogg-specific macro code

Sub GreenSheetFormat()

' Started by Paul Hively on 1/31/2017
' Automatic ColumnsToTabs for gift officer "Green Sheet" reports

' ************ EDIT BELOW HERE ************
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
    FieldsToKeep(16) = "Prefix"
    FieldsToKeep(17) = "FIRST"
    FieldsToKeep(18) = "Middle_Name"
    FieldsToKeep(19) = "LAST"
' ************ DO NOT EDIT BELOW HERE ************



End Sub