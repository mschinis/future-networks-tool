VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NetworkSpecifications 
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5430
   OleObjectBlob   =   "NetworkSpecifications.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NetworkSpecifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()
        TextBox3.MultiLine = True
       TextBox3.Text = "Load Density: 0.5MW/sqr km" & Chr(13) & "Number of Customers: 132" & Chr(13) & "Transformer Rating: 200kVA" & Chr(13) & "4 Feeders conductors: 185 mm sqr, 0.164+j0.069 ohms/km" & Chr(13) & "4 Laterals conductors: 95mm sqr, 0.320+j0.069 ohms/km"
End Sub

Private Sub CommandButton2_Click()
        TextBox2.MultiLine = True
       TextBox2.Text = "Load Density: 2MW/sqr km" & Chr(13) & "Number of Customers: 468" & Chr(13) & "Transformer Ratting: 500kVA" & Chr(13) & "4 Feeders conductors: 185 mm sqr, 0.164+j0.069 ohms/km" & Chr(13) & "4 Laterals conductors: 95mm sqr, 0.320+j0.069 ohms/km"
End Sub

Private Sub CommandButton4_Click()
        NetworkSpecifications.Hide
        ChooseNetwork.Show
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Networks_Change()

    'Label3.MultiLine = True
    If Networks.Value = "Urban" Then
        Label3.Caption = "5MW/sqr km" & Chr(13) & Chr(13) & "632" & Chr(13) & Chr(13) & "800 kVA" & Chr(13) & Chr(13) & "185 mm sqr Underground Cable, 0.164+j0.069 ohms/km" & Chr(13) & Chr(13) & "95 mm sqr Underground Cable, 0.320+j0.069 ohms/km"
    ElseIf Networks.Value = "SemiUrban" Then
        Label3.Caption = "2MW/sqr km" & Chr(13) & Chr(13) & "468" & Chr(13) & Chr(13) & "500 kVA" & Chr(13) & Chr(13) & "185 mm sqr Underground Cable, 0.164+j0.069 ohms/km" & Chr(13) & Chr(13) & "95 mm sqr Underground Cable, 0.320+j0.069 ohms/km"
    ElseIf Networks.Value = "Rural" Then
        Label3.Caption = "0.5MW/sqr km" & Chr(13) & Chr(13) & "132" & Chr(13) & Chr(13) & "200 kVA" & Chr(13) & Chr(13) & "185 mm sqr Overhead Line," & Chr(13) & "0.164+j0.069 ohms/km" & Chr(13) & Chr(13) & "95mm sqr Overhead Line," & Chr(13) & "0.320+j0.069 ohms/km"
    End If

End Sub

Private Sub TextBox1_Change()
    
End Sub

Private Sub CommandButton1_Click()
       TextBox1.MultiLine = True
       TextBox1.Text = "Load Density: 5MW/sqr km" & Chr(13) & "Number of Customers: 632" & Chr(13) & "Transformer Rating: 800kVA" & Chr(13) & "4 Feeders conductors: 185 mm sqr, 0.164+j0.069 ohms/km" & Chr(13) & "4 Laterals conductors: 95mm sqr, 0.320+j0.069 ohms/km"
End Sub


Private Sub UserForm_Initialize()

   Networks.List = ChooseNetwork.SelectNetwork.List


End Sub
