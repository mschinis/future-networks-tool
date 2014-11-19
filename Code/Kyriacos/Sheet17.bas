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

    If PresetNetwork.Network = "Urban" Then Call DrawUrban
    If PresetNetwork.Network = "Rural" Then Call DrawRural
    If PresetNetwork.Network = "SemiUrban" Then Call DrawSemiUrban
    
    Call CurrentOverload

End Sub

Private Sub ShowExtra_Click()
    If Sheet10.Visible <> xlSheetVisible Then
        ShowExtra.Caption = "Hide output tabs"
    Else
        ShowExtra.Caption = "Show output tabs"
    End If
    
    Sheet10.Visible = Not Sheet11.Visible
    Sheet11.Visible = Not Sheet11.Visible
    Sheet12.Visible = Not Sheet12.Visible
    Sheet13.Visible = Not Sheet13.Visible
    Sheet14.Visible = Not Sheet14.Visible
    Sheet7.Visible = Not Sheet7.Visible
    Sheet8.Visible = Not Sheet8.Visible
    Sheet9.Visible = Not Sheet9.Visible
    Sheet1.Visible = Not Sheet1.Visible
    Sheets("limits").Visible = Not Sheets("limits").Visible


End Sub
