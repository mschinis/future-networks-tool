VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Feeders 
   Caption         =   "UserForm1"
   ClientHeight    =   13140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   22590
   OleObjectBlob   =   "Feeders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Feeders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Feeder1Laterals_Change()




End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_initialize()

    For i = 1 To 10
        Me.MultiPage1.Pages("Feeder" & i).Visible = False
    Next

    For i = 1 To CustomNetworkForm.NoFeeders

        Me.MultiPage1.Pages("Feeder" & i).Visible = True
        
    Next
    
    For i = 1 To 5
    
        Feeder1Laterals.AddItem i
    
    Next

End Sub

