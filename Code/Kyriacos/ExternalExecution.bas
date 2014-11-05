Attribute VB_Name = "ExternalExecution"
Sub open_form()
 'Application.Visible = False
 'frmAddClient.Show vbModeless
End Sub
Private Sub cmdClose_Click()
    'Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'ThisWorkbook.Close SaveChanges:=True
    'Application.Visible = True
    'Application.Quit
End Sub
