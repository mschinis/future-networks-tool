VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

    If PresetNetwork.Network = "Urban" Then Call DrawNetworks.DrawUrban
    If PresetNetwork.Network = "Rural" Then Call DrawNetworks.DrawRural
    If PresetNetwork.Network = "SemiUrban" Then Call DrawNetworks.DrawSemiUrban
    
    Call CurrentOverload

End Sub

Private Sub CommandButton2_Click()
    SelectGraphsForm.Show
End Sub

Private Sub ShowExtra_Click()
    If Sheet10.Visible <> xlSheetVisible Then
        ShowExtra.Caption = "Hide output tabs"
    Else
        ShowExtra.Caption = "Show output tabs"
    End If
    
    Sheet16.Visible = Not Sheet16.Visible
    Sheet10.Visible = Not Sheet11.Visible
    Sheet11.Visible = Not Sheet11.Visible
    Sheet12.Visible = Not Sheet12.Visible
    Sheet13.Visible = Not Sheet13.Visible
    Sheet14.Visible = Not Sheet14.Visible
    Sheet7.Visible = Not Sheet7.Visible
    Sheet8.Visible = Not Sheet8.Visible
    Sheet9.Visible = Not Sheet9.Visible
    Sheet1.Visible = Not Sheet1.Visible
    Sheet3.Visible = Not Sheet3.Visible
    Sheet4.Visible = Not Sheet4.Visible
    Sheet5.Visible = Not Sheet5.Visible
    Sheet6.Visible = Not Sheet6.Visible
    Sheet20.Visible = Not Sheet20.Visible
    Sheet23.Visible = Not Sheet23.Visible
    Sheet24.Visible = Not Sheet24.Visible
    Sheet25.Visible = Not Sheet25.Visible
    
    Sheets("limits").Visible = Not Sheets("limits").Visible


End Sub
