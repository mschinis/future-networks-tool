VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Transformer 
   Caption         =   "UserForm2"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5070
   OleObjectBlob   =   "Transformer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Transformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Transformer.Hide
    
    CustomNetworkForm.TransformerLabel1 = "Capacity: " & TRSize & " KVA"
    CustomNetworkForm.TransformerLabel2 = "Impedance: " & TRImpedance & " p.u"
    
    
End Sub
