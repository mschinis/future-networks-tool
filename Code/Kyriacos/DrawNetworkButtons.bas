Attribute VB_Name = "DrawNetworkButtons"
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
(ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Private Const MAX_PATH As Long = 260

Public Sub Feeder1_Current()
    DrawNetworkGraphsForm.Show
    
    'MsgBox "Hello"
End Sub

'
'
'Private Sub Popup_Graph_UserForm()
'    Dim ws As Worksheet
'    Dim wsTemp As Worksheet
'    Dim rng As Range
'    Dim oChrt As ChartObject
'
'    '~~> Set the sheet where you have the charts data
'    Set ws = [Sheet1]
'
'    '~~> This is your charts range
'    Set rng = ws.Range("A1:B3")
'
'    '~~> Delete the temp sheeet if it is there
'    Application.DisplayAlerts = False
'    On Error Resume Next
'    ThisWorkbook.Sheets("TempOutput").Delete
'    On Error GoTo 0
'    Application.DisplayAlerts = True
'
'    Set wsTemp = Worksheets("Select Graphs")
'    Set oChrt = wsTemp.ChartObjects
'
'    '~~> Export the chart as bmp to the temp drive
'    oChrt.Chart.Export filename:=TempPath & "TempChart.bmp", Filtername:="Bmp"
'
'    '~~> Load the image to the image control
'    Me.Image1.Picture = LoadPicture(TempPath & "TempChart.bmp")
'
'    '~~> Delete the temp sheet
'    Application.DisplayAlerts = False
'    wsTemp.Delete
'    Application.DisplayAlerts = True
'
'    '~~> Kill the temp file
'    On Error Resume Next
'    Kill TempPath & "TempChart.bmp"
'    On Error GoTo 0
'End Sub
'
''~~> Function to get the user's temp path
'Function TempPath() As String
'    TempPath = String$(MAX_PATH, Chr$(0))
'    GetTempPath MAX_PATH, TempPath
'    TempPath = Replace(TempPath, Chr$(0), "")
'End Function
'
