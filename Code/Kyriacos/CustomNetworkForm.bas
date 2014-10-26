VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomNetworkForm 
   Caption         =   "UserForm1"
   ClientHeight    =   13440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   21510
   OleObjectBlob   =   "CustomNetworkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustomNetworkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Feeders.Show

End Sub

Private Sub Feeder1_Click()

    Feeders.MultiPage1.Value = 0
    Feeders.Show

End Sub

Private Sub Feeder2_Click()

    Feeders.MultiPage1.Value = 1
    Feeders.Show

End Sub
Private Sub Feeder3_Click()

    Feeders.MultiPage1.Value = 2
    Feeders.Show

End Sub
Private Sub Feeder4_Click()

    Feeders.MultiPage1.Value = 3
    Feeders.Show

End Sub
Private Sub Feeder5_Click()

    Feeders.MultiPage1.Value = 4
    Feeders.Show

End Sub
Private Sub Feeder6_Click()

    Feeders.MultiPage1.Value = 5
    Feeders.Show

End Sub
Private Sub Feeder7_Click()

    Feeders.MultiPage1.Value = 6
    Feeders.Show

End Sub
Private Sub Feeder8_Click()

    Feeders.MultiPage1.Value = 7
    Feeders.Show

End Sub
Private Sub Feeder9_Click()

    Feeders.MultiPage1.Value = 8
    Feeders.Show

End Sub
Private Sub Feeder10_Click()

    Feeders.MultiPage1.Value = 9
    Feeders.Show

End Sub

Private Sub Image1_Click()
    
    Transformer.Show
    
End Sub



Private Sub NoFeeders_Change()

    CommandButton1.Visible = False
    For i = 1 To 10
        Controls("Feeder" & i).Visible = False
    Next
    
    If NoFeeders.Value > 0 And NoFeeders <= 10 Then
    
    CommandButton1.Visible = True
    
    For i = 1 To NoFeeders.Value
    
        Me.Controls("Feeder" & i).Visible = True
        
    Next
    
    End If
    
End Sub

Private Sub UserForm_Initialize()

    CommandButton1.Visible = False

    For i = 1 To 10
    
        NoFeeders.AddItem i
        Controls("Feeder" & i).Visible = False
        
    Next

End Sub
