VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WelcomeScreen 
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3960
   OleObjectBlob   =   "WelcomeScreen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WelcomeScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    WelcomeScreen.Hide
    ChooseNetwork.Show


End Sub

Private Sub CommandButton2_Click()

    
End Sub

Private Sub UserForm_initialize()
    
    ChooseNetwork.finished = False

End Sub
