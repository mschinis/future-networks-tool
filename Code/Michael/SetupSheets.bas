Attribute VB_Name = "SetupSheets"
Public Sub setupAll()
    'Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    SetupSheets.setupGraphs
    SetupSheets.setupFeederSheets
    SetupSheets.setupFeederCurrentRollingAverages
    SetupSheets.setupCurrentRollingAverages
    SetupSheets.setupVoltageRollingAverages
    
    Application.DisplayAlerts = True
    'Application.ScreenUpdating = True
End Sub

Public Sub setupFeederSheets()
    Dim ws As Worksheet
    Dim feederSheetVoltageTitle As Range
    Dim feederSheetCurrentTitle As Range
    
    For i = 1 To SharedClass.Settings.feeders
        'Delete Start and End feeders
        deleteSheet ("Feeder" & i & "Start")
        deleteSheet ("Feeder" & i & "End")
        
        ' Setup the feeder start sheet
        Set ws = addSheet("Feeder" & i & "Start")
        Set ws = setupFeederStart(i, ws)
        
        ' Setup the feeder end sheet
        Set ws = addSheet("Feeder" & i & "End")
        Set ws = setupFeederEnd(i, ws)
    Next i
End Sub

Private Sub setupGraphs()
    Dim ws As Worksheet
    Dim cht As Chart
    
    For i = 1 To SharedClass.Settings.feeders
        deleteSheet ("Feeder " & i & " Output")
        
        Set ws = addSheet("Feeder " & i & " Output")
    Next i
End Sub
Private Function addChart(ByVal sheetName As String, location As Range, ByVal chtName As String, ByVal chtTitle As String, ByVal dataSetA As String, ByVal dataSetB As String, ByVal dataSetC As String)
    ' xAxisRng as string
    
    Dim sheet As Worksheet
    Dim cht As Chart
    Dim rng As Range
    
    Set sheet = Sheets(sheetName)
    Set cht = sheet.Shapes.addChart.Chart
    
    With cht
        ' Set Data
        .ChartType = xlXYScatterLinesNoMarkers
        .SeriesCollection.NewSeries
        .SeriesCollection(1).name = "Phase A"
        .SeriesCollection(1).XValues = "=Transformer!A3:A1442"
        .SeriesCollection(1).Values = Range(dataSetA)
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).name = "Phase B"
        .SeriesCollection(2).XValues = "=Transformer!A3:A1442"
        .SeriesCollection(2).Values = Range(dataSetB)
        
        .SeriesCollection.NewSeries
        .SeriesCollection(3).name = "Phase C"
        .SeriesCollection(3).XValues = "=Transformer!A3:A1442"
        .SeriesCollection(3).Values = Range(dataSetC)
        
        ' Location
        .ChartArea.Left = location.Left
        .ChartArea.Top = location.Left
        .ChartArea.Width = location.Width
        .ChartArea.Height = location.Height
        
        ' Titles
        .HasTitle = True
        .ChartTitle.Characters.Text = chtTitle
        .Parent.name = chtName
    End With
End Function

Public Sub setupFeederCurrentRollingAverages()
    Dim ws As Worksheet
    Dim currentCol As Integer
    Dim rollingAverageCount As Integer
    
    deleteSheet ("FeederCurrentRollingAverages")
    
    Set ws = addSheet("FeederCurrentRollingAverages")
    
    ' Settings
    ' rollingAverageCount: is the amount of values average will be
    ' currentCol: the first value will be placed at this column
    currentCol = 2
    rollingAverageCount = 60
    
    ' Create all side labels
    ws.Range(Cells(1385, currentCol - 1).Address()).Value = "Minimum:"
    ws.Range(Cells(1386, currentCol - 1).Address()).Value = "Maximum:"
    ws.Range(Cells(1389, currentCol - 1).Address()).Value = "Feeder Minimum:"
    ws.Range(Cells(1390, currentCol - 1).Address()).Value = "Feeder Maximum:"
    
    For i = 1 To SharedClass.Settings.feeders
        ws.Range(Cells(1, currentCol).Address(), Cells(1, currentCol + 2).Address()).Merge
        ws.Range(Cells(1, currentCol).Address()).HorizontalAlignment = xlCenter
        ws.Range(Cells(1, currentCol).Address()).Value = "Feeder " & i
        
        For j = currentCol To currentCol + 2
            phase = Chr(64 + (1 + j - currentCol))
            ws.Range(Cells(2, j).Address()).Value = "Phase " & phase
            
            ' Create the rolling average formulas
            ws.Range(Cells(3, j).Address()).Formula = "=AVERAGE(Feeder" & i & "Start!" & Cells(4, SharedClass.Settings.laterals * 3 * 2 + 2 + j - currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(rollingAverageCount + 3, SharedClass.Settings.laterals * 3 * 2 + 2 + j - currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
            ws.Range(Cells(3, j).Address()).AutoFill Destination:=Range(Cells(3, j).Address() & ":" & Cells(1383, j).Address()), Type:=xlFillSeries
            ' Calculate the minimum and maximum of the rolling average formulas
            ws.Range(Cells(1385, j).Address()).Formula = "=MIN(" & Cells(3, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(1383, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
            ws.Range(Cells(1386, j).Address()).Formula = "=MAX(" & Cells(3, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(1383, j).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
        Next j
        
        ' Merge the minimum and maximum cells and center
        ws.Range(Cells(1389, currentCol).Address(), Cells(1389, currentCol + 2)).Merge
        ws.Range(Cells(1390, currentCol).Address(), Cells(1390, currentCol + 2)).Merge
        ws.Range(Cells(1389, currentCol).Address()).HorizontalAlignment = xlCenter
        ws.Range(Cells(1390, currentCol).Address()).HorizontalAlignment = xlCenter
        
        ' Calculate the minimum and maximum of the three phases
        ws.Range(Cells(1389, currentCol).Address()).Formula = "=MIN(" & Cells(1385, currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(1385, currentCol + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")/Limits!E4"
        ws.Range(Cells(1390, currentCol).Address()).Formula = "=MAX(" & Cells(1386, currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(1386, currentCol + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")/Limits!E4"
        
        currentCol = currentCol + 3
    Next i
End Sub

Public Sub setupCurrentRollingAverages()
    Dim ws As Worksheet
    Dim currentCol As Integer
    Dim rollingAverageCount As Integer
    
    deleteSheet ("CurrentRollingAverages")
    
    Set ws = addSheet("CurrentRollingAverages")
    
    rollingAverageCount = 60
    currentCol = 2
    ' Create all side labels
    ws.Range(Cells(1386, currentCol - 1).Address()).Value = "Minimum:"
    ws.Range(Cells(1387, currentCol - 1).Address()).Value = "Maximum:"
    ws.Range(Cells(1391, currentCol - 1).Address()).Value = "Minimum:"
    ws.Range(Cells(1392, currentCol - 1).Address()).Value = "Maximum:"
    
    For i = 1 To SharedClass.Settings.feeders
        ws.Range(Cells(1, currentCol).Address(), Cells(1, currentCol + SharedClass.Settings.laterals * 3 - 1).Address()).Merge
        ws.Range(Cells(1, currentCol).Address()).HorizontalAlignment = xlCenter
        ws.Range(Cells(1, currentCol).Address()).Value = "Feeder " & i & " Start Currents"
        
        For j = 1 To SharedClass.Settings.laterals
            ws.Range(Cells(2, currentCol).Address(), Cells(2, currentCol + 2).Address()).Merge
            ws.Range(Cells(2, currentCol).Address()).HorizontalAlignment = xlCenter
            ws.Range(Cells(2, currentCol).Address()).Value = "Lateral " & j
            For k = currentCol To currentCol + 2
                phase = Chr(64 + (k - currentCol + 1))
                ws.Range(Cells(3, k).Address()).Value = "Phase " & phase
                
                ' Create the rolling average formulas
                ws.Range(Cells(4, k).Address()).Formula = "=AVERAGE(Feeder" & i & "Start!" & Cells(4, SharedClass.Settings.laterals * 3 + 2 + (j - 1) * 3 + k - currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(rollingAverageCount + 3, SharedClass.Settings.laterals * 3 + 2 + (j - 1) * 3 + k - currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
                ws.Range(Cells(4, k).Address()).AutoFill Destination:=Range(Cells(4, k).Address() & ":" & Cells(1384, k).Address()), Type:=xlFillSeries
                ' Calculate the minimum and maximum of the rolling average formulas
                ws.Range(Cells(1386, k).Address()).Formula = "=MIN(" & Cells(4, k).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(1384, k).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
                ws.Range(Cells(1387, k).Address()).Formula = "=MAX(" & Cells(4, k).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(1384, k).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
            Next k
            
            ' Merge the minimum and maximum cells and center
            ws.Range(Cells(1391, currentCol).Address(), Cells(1391, currentCol + 2)).Merge
            ws.Range(Cells(1392, currentCol).Address(), Cells(1392, currentCol + 2)).Merge
            ws.Range(Cells(1391, currentCol).Address()).HorizontalAlignment = xlCenter
            ws.Range(Cells(1392, currentCol).Address()).HorizontalAlignment = xlCenter
            ' Calculate the minimum and maximum of the three phases
            ws.Range(Cells(1391, currentCol).Address()).Formula = "=MIN(" & Cells(1386, currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(1386, currentCol + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")/Limits!D4"
            ws.Range(Cells(1392, currentCol).Address()).Formula = "=MAX(" & Cells(1387, currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(1387, currentCol + 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")/Limits!D4"
            
            currentCol = currentCol + 3
        Next j
    Next i
End Sub

Public Sub setupVoltageRollingAverages()
    Dim ws As Worksheet
    Dim currentCol As Integer
    Dim rollingAverageCount As Integer
    Dim cellType As String
    
    deleteSheet ("VoltageRollingAverages")
    
    Set ws = addSheet("VoltageRollingAverages")
    rollingAverageCount = 10
    currentCol = 3
    
    ' Setup BusBarVoltage phases
    ws.Range(Cells(1, currentCol).Address(), Cells(2, currentCol + 2).Address()).Merge
    ws.Range(Cells(1, currentCol).Address()).HorizontalAlignment = xlCenter
    ws.Range(Cells(1, currentCol).Address()).VerticalAlignment = xlCenter
    ws.Range(Cells(1, currentCol).Address()).Value = "Feeder " & i & " Bus Bar Voltage"
    For i = currentCol To currentCol + 2
        phase = Chr(64 + (i - currentCol + 1))
        ws.Range(Cells(3, i).Address()).Value = "Phase " & phase
        ws.Range(Cells(4, i).Address()).Formula = "=AVERAGE(Transformer!" & Cells(3, 3 + i - currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(rollingAverageCount + 2, 3 + i - currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
        ws.Range(Cells(4, i).Address()).AutoFill Destination:=Range(Cells(4, i).Address() & ":" & Cells(1434, i).Address()), Type:=xlFillSeries
                
    Next i
    currentCol = currentCol + 3
    
    ' Setup Feeder start voltages
    For z = 1 To 2
        If z = 1 Then cellType = "Start" Else cellType = "End"
        For i = 1 To SharedClass.Settings.feeders
            ws.Range(Cells(1, currentCol).Address(), Cells(1, currentCol + SharedClass.Settings.laterals * 3 - 1).Address()).Merge
            ws.Range(Cells(1, currentCol).Address()).HorizontalAlignment = xlCenter
            ws.Range(Cells(1, currentCol).Address()).Value = "Feeder " & i & " " & cellType & " Voltages"
            
            For j = 1 To SharedClass.Settings.laterals
                ws.Range(Cells(2, currentCol).Address(), Cells(2, currentCol + 2).Address()).Merge
                ws.Range(Cells(2, currentCol).Address()).HorizontalAlignment = xlCenter
                ws.Range(Cells(2, currentCol).Address()).Value = "Lateral " & j
                
                For k = currentCol To currentCol + 2
                    phase = Chr(64 + (k - currentCol + 1))
                    
                    ws.Range(Cells(3, k).Address()).Value = "Phase " & phase
                    ' Give the fields a name
                    ws.Range(Cells(4, k).Address() & ":" & Cells(1434, k).Address()).name = "Feeder" & i & "Lateral" & j & "VoltagePhase" & phase
                    ' Create the rolling averages formula
                    ws.Range(Cells(4, k).Address()).Formula = "=AVERAGE(Feeder" & i & cellType & "!" & Cells(4, 2 + (j - 1) * 3 + k - currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & Cells(rollingAverageCount + 3, 2 + (j - 1) * 3 + k - currentCol).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
                    ws.Range(Cells(4, k).Address()).AutoFill Destination:=Range(Cells(4, k).Address() & ":" & Cells(1434, k).Address()), Type:=xlFillSeries
                
                    
                Next k
                currentCol = currentCol + 3
            Next j
        Next i
    Next z
End Sub

Private Function setupFeederStart(ByVal feeder As Integer, ws As Worksheet) As Worksheet
    ws.Range(Cells(1, 2), Cells(1, 1 + SharedClass.Settings.laterals * 3)).Merge
    ws.Range(Cells(1, 2).Address()).HorizontalAlignment = xlCenter
    ws.Range(Cells(1, 2).Address()).Value = "Voltages"
    ' Create the currents label
    ws.Range(Cells(1, 2 + SharedClass.Settings.laterals * 3), Cells(1 + SharedClass.Settings.laterals * 3 * 2)).Merge
    ws.Range(Cells(1, 2 + SharedClass.Settings.laterals * 3).Address()).HorizontalAlignment = xlCenter
    ws.Cells(1, 2 + SharedClass.Settings.laterals * 3) = "Currents"
    
    ' Context of j,k:
    ' j = 1 is for the voltages
    ' j = 2 is for the currents
    ' k is the number of lateral
    Dim currentCol As Integer
    currentCol = 2
    For j = 1 To 2
        If j = 1 Then cellType = "Voltage" Else cellType = "Current"
        If j = 1 Then cellTypeShort = "V" Else cellTypeShort = "I"
        
        For k = 1 To SharedClass.Settings.laterals
            ws.Range(Cells(2, currentCol), Cells(2, currentCol + 2)).Merge
            ws.Range(Cells(2, currentCol).Address()).HorizontalAlignment = xlCenter
            ws.Range(Cells(2, currentCol).Address()).Value = "Lateral " & k
            
            dataSetBaseName = "Feeder" & feeder & "StartLateral" & k & cellType & "Phase"
            
            For z = currentCol To currentCol + 2
                phase = Chr(64 + (1 + z - currentCol))
                ws.Range(Cells(3, z).Address()).Value = "Phase " & phase
                ws.Range(Cells(4, z).Address() & ":" & Cells(1443, z).Address()).name = dataSetBaseName & phase
            Next z
            currentCol = currentCol + 3
            
            addChart "Feeder " & feeder & " Output", Range("A1:F20"), "Feeder" & feeder & "Lateral" & k & "Start" & cellTypeShort, "Lateral " & k & " Start " & cellType, dataSetBaseName & "A", dataSetBaseName & "B", dataSetBaseName & "C"
        Next k
    Next j
    ws.Range(Cells(2, currentCol), Cells(2, currentCol + 2)).Merge
    ws.Range(Cells(2, currentCol).Address()).HorizontalAlignment = xlCenter
    ws.Range(Cells(2, currentCol).Address()).Value = "Feeder " & feeder & " Start Current"
    
    dataSetBaseName = "Feeder" & feeder & "StartCurrentPhase"
    For z = currentCol To currentCol + 2
        phase = Chr(64 + (1 + z - currentCol))
        ws.Range(Cells(3, z).Address()).Value = "Phase " & phase
        ws.Range(Cells(4, z).Address() & ":" & Cells(1443, z).Address()).name = dataSetBaseName & phase
    Next z
    
     addChart "Feeder " & feeder & " Output", Range("A1:J20"), "Feeder" & feeder & "StartI", "Start Current", dataSetBaseName & "A", dataSetBaseName & "B", dataSetBaseName & "C"
End Function

Private Function setupFeederEnd(ByVal feeder As Integer, ws As Worksheet) As Worksheet
    ' Create the voltages label
    ws.Range(Cells(1, 2), Cells(1, 1 + SharedClass.Settings.laterals * 3)).Merge
    ws.Range(Cells(1, 2).Address()).HorizontalAlignment = xlCenter
    ws.Range(Cells(1, 2).Address()).Value = "Voltages"
    ' Create the currents label
    ws.Range(Cells(1, 2 + SharedClass.Settings.laterals * 3), Cells(1 + SharedClass.Settings.laterals * 3 * 2)).Merge
    ws.Range(Cells(1, 2 + SharedClass.Settings.laterals * 3).Address()).HorizontalAlignment = xlCenter
    ws.Cells(1, 2 + SharedClass.Settings.laterals * 3) = "Currents"
    
    ' Context of j,k:
    ' j = 1 is for the voltages
    ' j = 2 is for the currents
    ' k is the number of lateral
    currentCol = 2
    For j = 1 To 2
        If j = 1 Then cellType = "Voltage" Else cellType = "Current"
        If j = 1 Then cellTypeShort = "V" Else cellTypeShort = "I"
        
        For k = 1 To SharedClass.Settings.laterals
            ws.Range(Cells(2, currentCol), Cells(2, currentCol + 2)).Merge
            ws.Range(Cells(2, currentCol).Address()).HorizontalAlignment = xlCenter
            ws.Range(Cells(2, currentCol).Address()).Value = "Lateral " & k
            
            dataSetBaseName = "Feeder" & feeder & "EndLateral" & k & cellType & "Phase"
            
            For z = currentCol To currentCol + 2
                phase = Chr(64 + (1 + z - currentCol))
                ws.Range(Cells(3, z).Address()).Value = "Phase " & phase
                ws.Range(Cells(4, z).Address() & ":" & Cells(1443, z).Address()).name = dataSetBaseName & phase
            Next z
            
            addChart "Feeder " & feeder & " Output", Range("A1:F20"), "Feeder" & feeder & "Lateral" & k & "End" & cellTypeShort, "Lateral " & k & " End " & cellType, dataSetBaseName & "A", dataSetBaseName & "B", dataSetBaseName & "C"
            
            currentCol = currentCol + 3
        Next k
    Next j
    ws.Range(Cells(2, currentCol), Cells(2, currentCol + 2)).Merge
    ws.Range(Cells(2, currentCol).Address()).HorizontalAlignment = xlCenter
    ws.Range(Cells(2, currentCol).Address).Value = "Feeder " & feeder & " Start Current"
    For z = currentCol To currentCol + 2
        phase = Chr(64 + (1 + z - currentCol))
        ws.Range(Cells(3, z).Address()).Value = "Phase " & phase
    Next z
End Function

Private Function deleteSheet(ByVal name As String)
    If miscMacros.WorksheetExists(name) Then
        Sheets(name).Delete
    End If
End Function

Private Function addSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    With ThisWorkbook
        Set ws = .Worksheets.Add(After:=.Sheets(.Sheets.Count))
    End With
    ws.name = name
    Set addSheet = ws
End Function
