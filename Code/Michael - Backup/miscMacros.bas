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
    Dim i, j, counter, iextra As Long
    Dim s As String
    Dim FileNum As Long
    Dim rangex
    Dim Direc As String
    RunHours = Start.RunHours
    Dim character As String
    Dim iextrastr As String
    
    character = Chr(68)
    
    Dim Transformer() As Double
    ReDim Transformer(1 To RunHours, 1 To 3)
    ' Transformer Values
    Dim Values() As Double
    ReDim Values(1 To RunHours, 1 To 1)
    ' Voltage Lateral Values
    Dim VLValues() As Double
    ReDim VLValues(1 To RunHours, 1 To 3)
    ' Current Lateral Values
    Dim ILValues() As Double
    ReDim ILValues(1 To RunHours, 1 To 3)
    
    ' Voltage Feeder Values
    Dim VFValues() As Double
    ReDim VFValues(1 To RunHours, 1 To 3)
    ' Current Feeder Values
    Dim IFValues() As Double
    ReDim IFValues(1 To RunHours, 1 To 3)
    
    Dim VFeederStart() As Double
    ReDim VFeederStart(1 To RunHours + 1, 1 To 3)
    Dim IFeederStart() As Double
    ReDim IFeederStart(1 To RunHours + 1, 1 To 3)
    
    Dim VLateralStart() As Double
    ReDim VLateralStart(1 To RunHours + 1, 1 To 3)
    Dim ILateralStart() As Double
    ReDim ILateralStart(1 To RunHours + 1, 1 To 3)
    

    
    Direc = PresetNetwork.Network & "LVNetwork_" '
    
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
    
    ' >>>> time series results (P and Q) for GSP
    
    'using ParserX
    Set Parser = Nothing ' destroy old object should it already exist
    Set Parser = New ParserXControl.ParserX
    Parser.AutoIncrement = True
    FileNum = FreeFile
    
    ' Energy meter for transformer
    Open Direc & "EXP_METERS.CSV" For Input As #FileNum
    Line Input #FileNum, s
    Line Input #FileNum, s
    Parser.CmdString = s
    iextra = Parser.IntValue
    iextra = Parser.IntValue
    iextra = Parser.IntValue
    iextrastr = Parser.StrValue
    
    Dim kWh As Integer
    kWh = Parser.IntValue
    Dim kVarh As Integer
    kVarh = Parser.IntValue
    
    For i = 1 To 10
        iextra = Parser.IntValue
    Next
    
    Dim zoneLosseskWh As Integer
    zoneLosseskWh = Parser.IntValue ' Zone Losses kWh
    
    Dim zoneLosseskVarh As Integer
    zoneLosseskVarh = Parser.IntValue
    
    Set WorkingSheet = Worksheets("Results Summary")
'    WorkingSheet.Range("C3").Value = CheckValues.MinVoltage
'    WorkingSheet.Range("C4").Value = CheckValues.MaxVoltage
'    WorkingSheet.Range("C6").Value = CheckValues.MinCurrentUseFeeder
'    WorkingSheet.Range("C7").Value = CheckValues.MaxCurrentUseFeeder
'    WorkingSheet.Range("C9").Value = CheckValues.MinCurrentUseLateral
'    WorkingSheet.Range("C10").Value = CheckValues.MaxCurrentUseLateral
'    WorkingSheet.Range("C12").Value = CheckValues.MinTransformerUse
'    WorkingSheet.Range("C13").Value = CheckValues.MaxTransformerUse
    WorkingSheet.Range("C15").Value = ((zoneLosseskWh ^ 2 + zoneLosseskVarh ^ 2) ^ 0.5 / (kWh ^ 2 + kVarh ^ 2) ^ 0.5) 'Calculate losses %
    WorkingSheet.Range("C17").Value = CheckValues.VoltageCompliance
    WorkingSheet.Range("C18").Value = CheckValues.PercentageCustomersVoltage
    
    Close
    
    Set WorkingSheet = Worksheets("Transformer")
    Direc = Direc & "Mon_"
    ' Monitors for transformer
    i = 0
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
        If iextra > 90 Or iextra < -90 Then Transformer(i, 1) = -Transformer(i, 1)
        Transformer(i, 2) = Parser.DblValue
        iextra = Parser.DblValue
        If iextra > 90 Or iextra < -90 Then Transformer(i, 2) = -Transformer(i, 2)
        Transformer(i, 3) = Parser.DblValue
        iextra = Parser.DblValue
        If iextra > 90 Or iextra < -90 Then Transformer(i, 3) = -Transformer(i, 3)
        Values(i, 1) = Transformer(i, 1) + Transformer(i, 2) + Transformer(i, 3)

    Loop
    Close
    WorkingSheet.Range("B3:B" & (RunHours + 2)).Value = Values
    
    ' Feeders
    For i = 1 To 4
    counter = 0
    Set WorkingSheet = Worksheets("Feeder" & i & "Start")
        Open Direc & "vifeeder" & i & ".csv" For Input As #FileNum
        Line Input #FileNum, s
        Do While Not EOF(FileNum)
            Line Input #FileNum, s
            Parser.CmdString = s
            counter = counter + 1
            iextra = Parser.IntValue
            iextra = Parser.IntValue
            
                ' Voltages
                VFeederStart(counter, 1) = Parser.DblValue
                iextra = Parser.DblValue
                VFeederStart(counter, 2) = Parser.DblValue
                iextra = Parser.DblValue
                VFeederStart(counter, 3) = Parser.DblValue
                iextra = Parser.DblValue
                
                ' Currents
                IFeederStart(counter, 1) = Parser.DblValue
                iextra = Parser.DblValue
              
              
                If iextra > 40 Or iextra < -140 Then IFeederStart(counter, 1) = -IFeederStart(counter, 1)
                IFeederStart(counter, 2) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra > 100 Or iextra > -80 Then IFeederStart(counter, 2) = -IFeederStart(counter, 2)
                IFeederStart(counter, 3) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra < -5 Or iextra > 175 Then IFeederStart(counter, 3) = -IFeederStart(counter, 3)
                
                VFValues(counter, 1) = VFeederStart(counter, 1) / 230
                VFValues(counter, 2) = VFeederStart(counter, 2) / 230
                VFValues(counter, 3) = VFeederStart(counter, 3) / 230
                
                IFValues(counter, 1) = IFeederStart(counter, 1)
                IFValues(counter, 2) = IFeederStart(counter, 2)
                IFValues(counter, 3) = IFeederStart(counter, 3)
            
        Loop
        Close
        
        counter = 0
        character = Chr(65) ' Letter A
        ' Laterals
        For j = 1 To 4
            character = Chr(Asc(character) + 1)
            Open Direc & "vilateral" & i & "_" & j & "_start.csv" For Input As #FileNum
            counter = 0
            Line Input #FileNum, s
            Do While Not EOF(FileNum)
                Line Input #FileNum, s
                Parser.CmdString = s
                counter = counter + 1
                ' Skip hour and minute
                iextra = Parser.IntValue
                iextra = Parser.IntValue
                
                ' Voltages
                VLateralStart(counter, 1) = Parser.DblValue
                iextra = Parser.DblValue
                VLateralStart(counter, 2) = Parser.DblValue
                iextra = Parser.DblValue
                VLateralStart(counter, 3) = Parser.DblValue
                iextra = Parser.DblValue
                
                ' Currents
                ILateralStart(counter, 1) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra < 40 And iextra > -140 Then ILateralStart(counter, 1) = -ILateralStart(counter, 1)
                ILateralStart(counter, 2) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra > 100 Or iextra < -80 Then ILateralStart(counter, 2) = -ILateralStart(counter, 2)
                ILateralStart(counter, 3) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra > -20 And iextra < 160 Then ILateralStart(counter, 3) = -ILateralStart(counter, 3)
                
                VLValues(counter, 1) = VLateralStart(counter, 1) / 230
                VLValues(counter, 2) = VLateralStart(counter, 2) / 230
                VLValues(counter, 3) = VLateralStart(counter, 3) / 230
                
                ILValues(counter, 1) = ILateralStart(counter, 1)
                ILValues(counter, 2) = ILateralStart(counter, 2)
                ILValues(counter, 3) = ILateralStart(counter, 3)
            Loop
            Close
            ' Display Lateral Voltages
            WorkingSheet.Range(WorkingSheet.Cells(4, j * 3 - 1), WorkingSheet.Cells(RunHours + 3, j * 3 + 1)).Value = VLValues
            ' Display Lateral Currents
            WorkingSheet.Range(WorkingSheet.Cells(4, 12 + j * 3 - 1), WorkingSheet.Cells(RunHours + 3, 12 + j * 3 + 1)).Value = ILValues
            
            ' Display Feeder Currents
            WorkingSheet.Range(WorkingSheet.Cells(4, 26), WorkingSheet.Cells(RunHours + 3, 28)).Value = IFValues
       Next
       For j = 1 To 4
            Open Direc & "vilateral" & i & "_" & j & "_end.csv" For Input As #FileNum
            Set WorkingSheet = Worksheets("Feeder" & i & "End")
            counter = 0
            Line Input #FileNum, s
            Do While Not EOF(FileNum)
                Line Input #FileNum, s
                Parser.CmdString = s
                counter = counter + 1
                ' Skip hour and minute
                iextra = Parser.IntValue
                iextra = Parser.IntValue
                
                ' Voltages
                VLateralStart(counter, 1) = Parser.DblValue
                iextra = Parser.DblValue
                VLateralStart(counter, 2) = Parser.DblValue
                iextra = Parser.DblValue
                VLateralStart(counter, 3) = Parser.DblValue
                iextra = Parser.DblValue
                
                ' Currents
                ILateralStart(counter, 1) = Parser.DblValue
                iextra = Parser.DblValue
                ILateralStart(counter, 2) = Parser.DblValue
                iextra = Parser.DblValue
                ILateralStart(counter, 3) = Parser.DblValue
                iextra = Parser.DblValue
                
                VLValues(counter, 1) = VLateralStart(counter, 1) / 230
                VLValues(counter, 2) = VLateralStart(counter, 2) / 230
                VLValues(counter, 3) = VLateralStart(counter, 3) / 230
                
                ILValues(counter, 1) = ILateralStart(counter, 1)
                ILValues(counter, 2) = ILateralStart(counter, 2)
                ILValues(counter, 3) = ILateralStart(counter, 3)
            Loop
            Close
            ' Display Lateral Voltages
            WorkingSheet.Range(WorkingSheet.Cells(4, j * 3 - 1), WorkingSheet.Cells(RunHours + 3, j * 3 + 1)).Value = VLValues
            ' Display Lateral Currents
            WorkingSheet.Range(WorkingSheet.Cells(4, 12 + j * 3 - 1), WorkingSheet.Cells(RunHours + 3, 12 + j * 3 + 1)).Value = ILValues
        Next
        ' Display Feeders
    Next
    ' Display Feeder Voltages
    Set WorkingSheet = Worksheets("Transformer")
    WorkingSheet.Range(WorkingSheet.Cells(3, 3), WorkingSheet.Cells(RunHours + 2, 5)).Value = VFValues
    
    
    Set WorkingSheet = Worksheets("Limits")
    WorkingSheet.Range("D4:D1443").Value = CheckValues.lateralcurrentmax
    'workingsheet.Range("E4:E1443").Value = -CheckValues.lateralcurrentmax
    WorkingSheet.Range("E4:F1443").Value = CheckValues.feedercurrentmax
    WorkingSheet.Range("F4:F1443").Value = -CheckValues.feedercurrentmax
    WorkingSheet.Range("G4:G1443").Value = CheckValues.TransformerMax
    WorkingSheet.Range("H4:H1443").Value = -CheckValues.TransformerMax
    
End Sub

Public Sub Customers_Voltage()

    Worksheets("test").Range("A1:CCC632").Clear
    Dim i, y, iter, place As Integer
    
    For i = 1 To 4
        For y = 1 To PresetNetwork.customers / 4
            For iter = 1 To Start.RunHours
                place = (i * (PresetNetwork.customers / 4) - (PresetNetwork.customers / 4)) + y
                Worksheets("Test").Cells(place, iter).Value = Start.CustomersLimits(i, y, iter)
            Next
        Next
    Next
    Dim cs As ColorScale
    
    Worksheets("Test").Activate
    With Worksheets("Test").Range(Cells(1, 1), Cells(PresetNetwork.customers, Start.RunHours))
        .FormatConditions.Delete
        Set cs = .FormatConditions.AddColorScale(colorscaletype:=2)
        With cs.ColorScaleCriteria(1)
            .Type = xlConditionValueNumber
            .Value = 0
            With .FormatColor
                .Color = vbGreen
                ' TintAndShade takes a value between -1 and 1.
                ' -1 is darkest, 1 is lightest.
                .TintAndShade = -0.25
            End With
        End With
   
        ' Format the second color as green, at the highest value.
        With cs.ColorScaleCriteria(2)
            .Type = xlConditionValueNumber
            .Value = 1
            With .FormatColor
                .Color = vbRed
                .TintAndShade = 0
            End With
        End With
    End With
    
End Sub


