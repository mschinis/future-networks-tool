VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ShowExtra_Click()
    
    Sheet10.Visible = True
    Sheet11.Visible = True
    Sheet12.Visible = True
    Sheet13.Visible = True
    Sheet14.Visible = True
    Sheet7.Visible = True
    Sheet8.Visible = True
    Sheet9.Visible = True
    Sheet1.Visible = True
    HideExtra.Visible = True
    ShowExtra.Visible = False


End Sub
Private Sub HideExtra_Click()
    
    Sheet10.Visible = False
    Sheet11.Visible = False
    Sheet12.Visible = False
    Sheet13.Visible = False
    Sheet14.Visible = False
    Sheet7.Visible = False
    Sheet8.Visible = False
    Sheet9.Visible = False
    Sheet1.Visible = False
    ShowExtra.Visible = True
    HideExtra.Visible = False


End Sub
