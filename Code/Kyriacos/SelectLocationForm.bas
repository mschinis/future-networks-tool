VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectLocationForm 
   Caption         =   "UserForm1"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8535
   OleObjectBlob   =   "SelectLocationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectLocationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Unload Me

End Sub


Private Sub UserForm_Initialize()
    
    If ChooseNetwork.SelectLocation.ListIndex = 0 Then Scotland.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 1 Then NorthEast.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 2 Then NorthWest.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 3 Then York.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 4 Then EastMidlands.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 5 Then WestMidlands.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 6 Then East.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 7 Then Wales.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 8 Then London.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 9 Then SouthEast.Value = True
    If ChooseNetwork.SelectLocation.ListIndex = 10 Then SouthWest.Value = True

End Sub

Private Sub UserForm_Terminate()
    If Scotland.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 0
    If NorthEast.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 1
    If NorthWest.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 2
    If York.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 3
    If EastMidlands.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 4
    If WestMidlands.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 5
    If East.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 6
    If Wales.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 7
    If London.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 8
    If SouthEast.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 9
    If SouthWest.Value = True Then ChooseNetwork.SelectLocation.ListIndex = 10
     
    If ChooseNetwork.SelectLocation.Value = "" Then
        MsgBox "No location selected"
    Else
        MsgBox ChooseNetwork.SelectLocation.Value & " selected"
    End If

End Sub
