VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdvancedProperties 
   Caption         =   "Advanced Settings"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   OleObjectBlob   =   "AdvancedProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdvancedProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    Call UserForm_Initialize


End Sub

Private Sub CommandButton2_Click()
    
If Start.OverrideDefault = True Then
    
    MSG1 = MsgBox("Changing these values might affect the way the network operates. Proceed anyway?", vbYesNo, "Warning!")

    If MSG1 = vbYes Then
        AdvancedProperties.Hide
        
    Else
        
    End If
Else
    AdvancedProperties.Hide
End If


End Sub

 
Private Sub FeederMax_Change()
Start.OverrideDefault = True
End Sub



Private Sub Label10_Click()

End Sub

Private Sub LateralMax_Change()
Start.OverrideDefault = True
End Sub

Private Sub TransformerMax_Change()
Start.OverrideDefault = True
End Sub

Private Sub TransformerVoltage_Change()
Start.OverrideDefault = True
End Sub

Private Sub UserForm_Initialize()

If Start.OverrideDefault = False Then

    TransformerVoltage.Value = 433
    VoltageMin.Value = 0.9
    VoltageMax.Value = 1.1
    VoltageAverageMin.Value = 0.94
    TransformerMax.Value = 100
    FeederMax.Value = 100
    LateralMax.Value = 100
    
''    If ChooseNetwork.SelectNetwork = "Urban" Then
''        TransformerMax.Value = 800
''        If ChooseNetwork.TdayVal.Value <= 4 Or ChooseNetwork.TdayVal.Value >= 11 Then
''            FeederMax.Value = 309
''            LateralMax.Value = 209
''        Else
''            FeederMax.Value = 297
''            LateralMax.Value = 202
''        End If
''
''    ElseIf ChooseNetwork.SelectNetwork = "SemiUrban" Then
''        TransformerMax.Value = 500
''        If ChooseNetwork.TdayVal.Value <= 4 Or ChooseNetwork.TdayVal.Value >= 11 Then
''            FeederMax.Value = 309
''            LateralMax.Value = 209
''        Else
''            FeederMax.Value = 297
''            LateralMax.Value = 202
''        End If
''
''    ElseIf ChooseNetwork.SelectNetwork = "Rural" Then
''        TransformerMax.Value = 200
''        If ChooseNetwork.TdayVal.Value <= 4 Or ChooseNetwork.TdayVal.Value >= 11 Then
''            FeederMax.Value = 404
''            LateralMax.Value = 263
''        Else
''            FeederMax.Value = 350
''            LateralMax.Value = 230
''        End If
''    End If
End If




Start.OverrideDefault = False
    
End Sub

Private Sub VoltageAverageMin_Change()
Start.OverrideDefault = True
End Sub

Private Sub VoltageMax_Change()
Start.OverrideDefault = True
End Sub

Private Sub VoltageMin_Change()
Start.OverrideDefault = True
End Sub
