Attribute VB_Name = "miscMacros"
Option Explicit

Public DSSobj As OpenDSSengine.DSS
Public DSSText As OpenDSSengine.Text
Public DSSCircuit As OpenDSSengine.Circuit
Public DSSSolution As OpenDSSengine.Solution
Public DSSControlQueue As OpenDSSengine.CtrlQueue
Public Parser As ParserXControl.ParserX

Public Function File_Exists(ByVal File As String) As Boolean

    Dim FilePath As String
    Dim TestStr As String

    FilePath = ActiveWorkbook.Path & "\" & File
    
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        File_Exists = False
    Else
        File_Exists = True
    End If

End Function

Public Sub Monitors()

    Dim WorkingSheet As Worksheet
    Dim i, j, iextra As Long
    Dim s As String
    Dim FileNum As Long
    Dim rangex
    Dim Direc As String
    RunHours = Start.RunHours
    
    Dim Transformer() As Double
    ReDim Transformer(1 To RunHours, 1 To 3)
    Dim Values() As Double
    ReDim Values(1 To RunHours, 1 To 1)
    
    Direc = PresetNetwork.Network & "LVNetwork_Mon_" '
    
    ' Export dem monitors
    DSSText.Command = "Export monitors SSTransformer"
    
    For i = 1 To 4
        DSSText.Command = "Export monitors VIFeeder" & i
        
        For j = 1 To 4
            DSSText.Command = "Export monitors VILateral" & i & "_" & j & "_Start"
            DSSText.Command = "Export monitors VILateral" & i & "_" & j & "_End"
        Next
    Next
    ' TODO: Check if necessary files exist
    ' Start of feeder
    ' Transformer
    ' Start, End of each lateral
    
    
    ' >>>>
    ' >>>> time series results (P and Q) for GSP
    Set WorkingSheet = Worksheets("Transformer")
    'using ParserX
    Set Parser = Nothing ' destroy old object should it already exist
    Set Parser = New ParserXControl.ParserX
    Parser.AutoIncrement = True
    FileNum = FreeFile
    i = 0
    ' TODO: Check if necessary files exist
    Open Direc & "transformer.csv" For Input As #FileNum
    Line Input #FileNum, s  ' skip first line
    Do While Not EOF(FileNum)
        Line Input #FileNum, s
        Parser.CmdString = s
        i = i + 1
        iextra = Parser.IntValue 'hours
        iextra = Parser.IntValue 'seconds
        Transformer(i, 1) = Parser.DblValue
        iextra = Parser.DblValue
        Transformer(i, 2) = Parser.DblValue
        iextra = Parser.DblValue
        Transformer(i, 3) = Parser.DblValue
        iextra = Parser.DblValue
        Values(i, 1) = Transformer(i, 1) + Transformer(i, 2) + Transformer(i, 3)

    Loop
    
    WorkingSheet.Range("B2:B" & (RunHours + 1)).Value = Values
    
End Sub


