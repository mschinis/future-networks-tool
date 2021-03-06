Attribute VB_Name = "miscMacros"
Option Explicit

Public DSSobj As OpenDSSengine.DSS
Public DSSText As OpenDSSengine.Text
Public DSSCircuit As OpenDSSengine.Circuit
Public DSSSolution As OpenDSSengine.Solution
Public DSSControlQueue As OpenDSSengine.CtrlQueue
Public Parser As ParserXControl.ParserX

Public Function WorksheetExists(ByVal WorksheetName As String) As Boolean

    On Error Resume Next
    WorksheetExists = (Sheets(WorksheetName).name <> "")
    On Error GoTo 0

End Function

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
Public Sub CheckDir(strDirectoryPath As String)

If Dir(strDirectoryPath) = "" Then
    MkDir strDirectoryPath
End If

End Sub
Public Sub Monitors()

    Dim workingsheet As Worksheet
    Dim i, j, counter, iextra As Long
    Dim s As String
    Dim FileNum As Long
    Dim rangex
    Dim Direc As String
    RunHours = Start.RunHours
    Dim character As String
    Dim iextrastr As String
    Dim z As Integer
    Dim countercorrected As Integer
    Dim AngleVariable As Integer
    
    AngleVariable = 30
    
    
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
    

    
    Direc = PresetNetwork.network & "LVNetwork_" '
    
    ' Export dem monitors
    DSSText.Command = "Export monitors SSTransformer"
    
    For i = 1 To SharedClass.Settings.feeders
        DSSText.Command = "Export monitors VIFeeder" & i
 '       DSSText.Command = "Export monitors VIFeeder" & i & "n"
        
        For j = 1 To SharedClass.Settings.laterals
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
    
    Set workingsheet = Worksheets("Results Summary")
'    WorkingSheet.Range("C3").Value = CheckValues.MinVoltage
'    WorkingSheet.Range("C4").Value = CheckValues.MaxVoltage
'    WorkingSheet.Range("C6").Value = CheckValues.MinCurrentUseFeeder
'    WorkingSheet.Range("C7").Value = CheckValues.MaxCurrentUseFeeder
'    WorkingSheet.Range("C9").Value = CheckValues.MinCurrentUseLateral
'    WorkingSheet.Range("C10").Value = CheckValues.MaxCurrentUseLateral
'    WorkingSheet.Range("C12").Value = CheckValues.MinTransformerUse
'    WorkingSheet.Range("C13").Value = CheckValues.MaxTransformerUse
    workingsheet.Range("C15").Value = ((zoneLosseskWh ^ 2 + zoneLosseskVarh ^ 2) ^ 0.5 / (kWh ^ 2 + kVarh ^ 2) ^ 0.5) 'Calculate losses %
    workingsheet.Range("C17").Value = CheckValues.VoltageCompliance
    workingsheet.Range("C18").Value = CheckValues.PercentageCustomersVoltage
    
    Close
    
    Set workingsheet = Worksheets("Transformer")
    Direc = Direc & "Mon_"
    ' Monitors for transformer
    i = 0
    Open Direc & "transformer.csv" For Input As #FileNum
    For z = 1 To 421
        Line Input #FileNum, s  ' skip first 7 hours
    Next
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
        
        If i <= 1020 Then
            Values(i + 420, 1) = Transformer(i, 1) + Transformer(i, 2) + Transformer(i, 3)
        Else
            Values(i - 1020, 1) = Transformer(i, 1) + Transformer(i, 2) + Transformer(i, 3)
        End If

    Loop
    Close
    workingsheet.Range("B3:B" & (RunHours + 2)).Value = Values
    
    ' Feeders
    For i = 1 To SharedClass.Settings.feeders
        counter = 0
        Set workingsheet = Worksheets("Feeder" & i & "Start")
        Open Direc & "vifeeder" & i & ".csv" For Input As #FileNum
        For z = 1 To 421
            Line Input #FileNum, s  ' skip first 7 hours
        Next
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
                If iextra < 30 + AngleVariable And iextra > -150 + AngleVariable Then IFeederStart(counter, 1) = -IFeederStart(counter, 1)
                
                IFeederStart(counter, 2) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra > 90 + AngleVariable Or iextra < -90 + AngleVariable Then IFeederStart(counter, 2) = -IFeederStart(counter, 2)
                
                IFeederStart(counter, 3) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra > -30 + AngleVariable And iextra < 150 + AngleVariable Then IFeederStart(counter, 3) = -IFeederStart(counter, 3)
                
                If counter <= 1020 Then
                    countercorrected = counter + 420
                Else
                    countercorrected = counter - 1020
                End If
                
                
                VFValues(countercorrected, 1) = VFeederStart(counter, 1) / 230
                VFValues(countercorrected, 2) = VFeederStart(counter, 2) / 230
                VFValues(countercorrected, 3) = VFeederStart(counter, 3) / 230
                
                IFValues(countercorrected, 1) = IFeederStart(counter, 1)
                IFValues(countercorrected, 2) = IFeederStart(counter, 2)
                IFValues(countercorrected, 3) = IFeederStart(counter, 3)
            
        Loop
        Close
        
        counter = 0
        character = Chr(65) ' Letter A
        ' Laterals
        For j = 1 To SharedClass.Settings.laterals
            character = Chr(Asc(character) + 1)
            Open Direc & "vilateral" & i & "_" & j & "_start.csv" For Input As #FileNum
            counter = 0
            For z = 1 To 421
                Line Input #FileNum, s  ' skip first 7 hours
            Next z
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
                If iextra < 30 + AngleVariable And iextra > -150 + AngleVariable Then ILateralStart(counter, 1) = -ILateralStart(counter, 1)
                
                ILateralStart(counter, 2) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra > 90 + AngleVariable Or iextra < -90 + AngleVariable Then ILateralStart(counter, 2) = -ILateralStart(counter, 2)
                
                ILateralStart(counter, 3) = Parser.DblValue
                iextra = Parser.DblValue
                If iextra > -30 + AngleVariable And iextra < 150 + AngleVariable Then ILateralStart(counter, 3) = -ILateralStart(counter, 3)
                
                If counter <= 1020 Then
                    countercorrected = counter + 420
                Else
                    countercorrected = counter - 1020
                End If
                
                VLValues(countercorrected, 1) = VLateralStart(counter, 1) / 230
                VLValues(countercorrected, 2) = VLateralStart(counter, 2) / 230
                VLValues(countercorrected, 3) = VLateralStart(counter, 3) / 230
                
                ILValues(countercorrected, 1) = ILateralStart(counter, 1)
                ILValues(countercorrected, 2) = ILateralStart(counter, 2)
                ILValues(countercorrected, 3) = ILateralStart(counter, 3)
            Loop
            Close
            ' Display Lateral Voltages
            workingsheet.Range(workingsheet.Cells(4, j * 3 - 1), workingsheet.Cells(RunHours + 3, j * 3 + 1)).Value = VLValues
            ' Display Lateral Currents
            workingsheet.Range(workingsheet.Cells(4, 12 + j * 3 - 1), workingsheet.Cells(RunHours + 3, 12 + j * 3 + 1)).Value = ILValues
            
            ' Display Feeder Currents
            workingsheet.Range(workingsheet.Cells(4, 26), workingsheet.Cells(RunHours + 3, 28)).Value = IFValues
        Next j
        For j = 1 To SharedClass.Settings.laterals
            Open Direc & "vilateral" & i & "_" & j & "_end.csv" For Input As #FileNum
            Set workingsheet = Worksheets("Feeder" & i & "End")
            counter = 0
            For z = 1 To 421
                Line Input #FileNum, s  ' skip first 7 hours
            Next
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
                
                If counter <= 1020 Then
                    countercorrected = counter + 420
                Else
                    countercorrected = counter - 1020
                End If
                
                VLValues(countercorrected, 1) = VLateralStart(counter, 1) / 230
                VLValues(countercorrected, 2) = VLateralStart(counter, 2) / 230
                VLValues(countercorrected, 3) = VLateralStart(counter, 3) / 230
                
                ILValues(countercorrected, 1) = ILateralStart(counter, 1)
                ILValues(countercorrected, 2) = ILateralStart(counter, 2)
                ILValues(countercorrected, 3) = ILateralStart(counter, 3)
            Loop
            Close
            ' Display Lateral Voltages
            workingsheet.Range(workingsheet.Cells(4, j * 3 - 1), workingsheet.Cells(RunHours + 3, j * 3 + 1)).Value = VLValues
            ' Display Lateral Currents
            workingsheet.Range(workingsheet.Cells(4, 12 + j * 3 - 1), workingsheet.Cells(RunHours + 3, 12 + j * 3 + 1)).Value = ILValues
        Next
        ' Display Feeders
    Next
    ' Display Feeder Voltages
    Set workingsheet = Worksheets("Transformer")
    workingsheet.Range(workingsheet.Cells(3, 3), workingsheet.Cells(RunHours + 2, 5)).Value = VFValues
    
    
    Set workingsheet = Worksheets("Limits")
    workingsheet.Range("D4:D1443").Value = CheckValues.lateralcurrentmax
    'workingsheet.Range("E4:E1443").Value = -CheckValues.lateralcurrentmax
    workingsheet.Range("E4:F1443").Value = CheckValues.feedercurrentmax
    workingsheet.Range("F4:F1443").Value = -CheckValues.feedercurrentmax
    workingsheet.Range("G4:G1443").Value = CheckValues.TransformerMax
    workingsheet.Range("H4:H1443").Value = -CheckValues.TransformerMax
    
End Sub

Public Sub Customers_Voltage()

    Worksheets("test").Range("A1:CCC632").Clear
    Dim i, y, iter, place As Integer
    
    For i = 1 To SharedClass.Settings.feeders
        feedercustomers = 0
        For m = 1 To SharedClass.Settings.laterals
            feedercustomers = feedercustomers + Assign_Profiles.LateralSizes(i, m)
        Next
        
        For y = 1 To feedercustomers
            For iter = 1 To Start.RunHours
                place = (i * (feedercustomers) - (feedercustomers)) + y
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


