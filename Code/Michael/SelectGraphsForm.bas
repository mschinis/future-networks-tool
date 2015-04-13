VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectGraphsForm 
   Caption         =   "Select graphs"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11520
   OleObjectBlob   =   "SelectGraphsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectGraphsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim workingsheet As Worksheet

Private Sub SelectGraphsButtonPressed_Click()
    Dim optionName As String
    Dim noSelectedGraphs As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tempInt As Integer
    Dim currentGraphName As String
    Dim graph As ChartObject
    Dim yLocation As Integer
    j = 0
    
    ' Clear graphs if any exist inside "Select Graphs" Worksheet
    Set workingsheet = Worksheets("Select Graphs")
    If workingsheet.ChartObjects.Count > 0 Then
        workingsheet.ChartObjects.Delete
    End If
    
    ' Dat logic tho
    For i = 0 To SelectGraphsList.ListCount - 1
        Set workingsheet = Sheets("Transformer Output")
        If SelectGraphsList.Selected(i) Then
            Select Case SelectGraphsList.List(i)
                Case "Transformer Power Output"
                    currentGraphName = "TransformerPowerOutputGraph"
                Case "BusBar Voltage"
                    currentGraphName = "BusBarVoltageGraph"
                Case Else
                
            End Select
            
            ' Select the graph, copy it, change workingsheet
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartArea.Copy
            Set workingsheet = Sheets("Select Graphs")
            ' Determine the position of the new graph
            tempInt = j * 20
            yLocation = j * 20 + 19
            If tempInt = 0 Then tempInt = 1
            ' Paste it, and change the title
            workingsheet.Paste workingsheet.Range("D" & tempInt)
            Set graph = workingsheet.ChartObjects(currentGraphName)
            
            Set RngToCover = workingsheet.Range("D" & tempInt & ":O" & yLocation)
            With graph
                .Height = RngToCover.Height
                .Width = RngToCover.Width
                .Top = RngToCover.Top
                .Left = RngToCover.Left
            End With
            j = j + 1
        End If
    Next i
    
    ' Feeder 1
    For i = 0 To SelectGraphsFeederOneList.ListCount - 1
        Set workingsheet = Sheets("Feeder 1 Output")
        If SelectGraphsFeederOneList.Selected(i) Then
            Select Case SelectGraphsFeederOneList.List(i)
                Case "Lateral 1 Start Voltage"
                    currentGraphName = "Feeder1Lateral1StartV"
                    
                Case "Lateral 2 Start Voltage"
                    currentGraphName = "Feeder1Lateral2StartV"
                Case "Lateral 3 Start Voltage"
                    currentGraphName = "Feeder1Lateral3StartV"
                Case "Lateral 4 Start Voltage"
                    currentGraphName = "Feeder1Lateral4StartV"
                Case "Lateral 1 End Voltage"
                    currentGraphName = "Feeder1Lateral1EndV"
                Case "Lateral 2 End Voltage"
                    currentGraphName = "Feeder1Lateral2EndV"
                Case "Lateral 3 End Voltage"
                    currentGraphName = "Feeder1Lateral3EndV"
                Case "Lateral 4 End Voltage"
                    currentGraphName = "Feeder1Lateral4EndV"

                Case "Feeder 1 Start Current"
                    currentGraphName = "Feeder1StartI"
                    
                Case "Lateral 1 Start Current"
                    currentGraphName = "Feeder1Lateral1StartI"
                Case "Lateral 2 Start Current"
                    currentGraphName = "Feeder1Lateral2StartI"
                Case "Lateral 3 Start Current"
                    currentGraphName = "Feeder1Lateral3StartI"
                Case "Lateral 4 Start Current"
                    currentGraphName = "Feeder1Lateral4StartI"
            End Select
            
            ' Select the graph, copy it, change workingsheet
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartArea.Copy
            Set workingsheet = Sheets("Select Graphs")
            ' Determine the position of the new graph
            tempInt = j * 20
            yLocation = j * 20 + 19
            If tempInt = 0 Then tempInt = 1
            ' Paste it, and change the title
            workingsheet.Paste workingsheet.Range("D" & tempInt)
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartTitle.Text = "Feeder 1 " & graph.Chart.ChartTitle.Text
            
            Set RngToCover = workingsheet.Range("D" & tempInt & ":O" & yLocation)
            With graph
                .Height = RngToCover.Height
                .Width = RngToCover.Width
                .Top = RngToCover.Top
                .Left = RngToCover.Left
            End With
            j = j + 1
        End If
    Next i
    
    ' Feeder 2
    For i = 0 To SelectGraphsFeederTwoList.ListCount - 1
        Set workingsheet = Sheets("Feeder 2 Output")
        If SelectGraphsFeederTwoList.Selected(i) Then
            Select Case SelectGraphsFeederTwoList.List(i)
                Case "Lateral 1 Start Voltage"
                    currentGraphName = "Feeder2Lateral1StartV"
                    
                Case "Lateral 2 Start Voltage"
                    currentGraphName = "Feeder2Lateral2StartV"
                Case "Lateral 3 Start Voltage"
                    currentGraphName = "Feeder2Lateral3StartV"
                Case "Lateral 4 Start Voltage"
                    currentGraphName = "Feeder2Lateral4StartV"
                Case "Lateral 1 End Voltage"
                    currentGraphName = "Feeder2Lateral1EndV"
                Case "Lateral 2 End Voltage"
                    currentGraphName = "Feeder2Lateral2EndV"
                Case "Lateral 3 End Voltage"
                    currentGraphName = "Feeder2Lateral3EndV"
                Case "Lateral 4 End Voltage"
                    currentGraphName = "Feeder2Lateral4EndV"

                Case "Feeder 2 Start Current"
                    currentGraphName = "Feeder2StartI"
                    
                Case "Lateral 1 Start Current"
                    currentGraphName = "Feeder2Lateral1StartI"
                Case "Lateral 2 Start Current"
                    currentGraphName = "Feeder2Lateral2StartI"
                Case "Lateral 3 Start Current"
                    currentGraphName = "Feeder2Lateral3StartI"
                Case "Lateral 4 Start Current"
                    currentGraphName = "Feeder2Lateral4StartI"
            End Select
            
            ' Select the graph, copy it, change workingsheet
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartArea.Copy
            Set workingsheet = Sheets("Select Graphs")
            ' Determine the position of the new graph
            tempInt = j * 20
            yLocation = j * 20 + 19
            If tempInt = 0 Then tempInt = 1
            ' Paste it, and change the title
            workingsheet.Paste workingsheet.Range("D" & tempInt)
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartTitle.Text = "Feeder 2 " & graph.Chart.ChartTitle.Text
            
            Set RngToCover = workingsheet.Range("D" & tempInt & ":O" & yLocation)
            With graph
                .Height = RngToCover.Height
                .Width = RngToCover.Width
                .Top = RngToCover.Top
                .Left = RngToCover.Left
            End With
            j = j + 1
        End If
    Next i
    
    
    ' Feeder 3
    For i = 0 To SelectGraphsFeederThreeList.ListCount - 1
        Set workingsheet = Sheets("Feeder 3 Output")
        If SelectGraphsFeederThreeList.Selected(i) Then
            Select Case SelectGraphsFeederThreeList.List(i)
                Case "Lateral 1 Start Voltage"
                    currentGraphName = "Feeder3Lateral1StartV"
                    
                Case "Lateral 2 Start Voltage"
                    currentGraphName = "Feeder3Lateral2StartV"
                Case "Lateral 3 Start Voltage"
                    currentGraphName = "Feeder3Lateral3StartV"
                Case "Lateral 4 Start Voltage"
                    currentGraphName = "Feeder3Lateral4StartV"
                Case "Lateral 1 End Voltage"
                    currentGraphName = "Feeder3Lateral1EndV"
                Case "Lateral 2 End Voltage"
                    currentGraphName = "Feeder3Lateral2EndV"
                Case "Lateral 3 End Voltage"
                    currentGraphName = "Feeder3Lateral3EndV"
                Case "Lateral 4 End Voltage"
                    currentGraphName = "Feeder3Lateral4EndV"

                Case "Feeder 3 Start Current"
                    currentGraphName = "Feeder3StartI"
                    
                Case "Lateral 1 Start Current"
                    currentGraphName = "Feeder3Lateral1StartI"
                Case "Lateral 2 Start Current"
                    currentGraphName = "Feeder3Lateral2StartI"
                Case "Lateral 3 Start Current"
                    currentGraphName = "Feeder3Lateral3StartI"
                Case "Lateral 4 Start Current"
                    currentGraphName = "Feeder3Lateral4StartI"
            End Select
            
            ' Select the graph, copy it, change workingsheet
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartArea.Copy
            Set workingsheet = Sheets("Select Graphs")
            ' Determine the position of the new graph
            tempInt = j * 20
            yLocation = j * 20 + 19
            If tempInt = 0 Then tempInt = 1
            ' Paste it, and change the title
            workingsheet.Paste workingsheet.Range("D" & tempInt)
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartTitle.Text = "Feeder 3 " & graph.Chart.ChartTitle.Text
            
            Set RngToCover = workingsheet.Range("D" & tempInt & ":O" & yLocation)
            With graph
                .Height = RngToCover.Height
                .Width = RngToCover.Width
                .Top = RngToCover.Top
                .Left = RngToCover.Left
            End With
            j = j + 1
        End If
    Next i
    
    
    ' Feeder 4
    For i = 0 To SelectGraphsFeederFourList.ListCount - 1
        Set workingsheet = Sheets("Feeder 4 Output")
        If SelectGraphsFeederFourList.Selected(i) Then
            Select Case SelectGraphsFeederFourList.List(i)
                Case "Lateral 1 Start Voltage"
                    currentGraphName = "Feeder4Lateral1StartV"
                    
                Case "Lateral 2 Start Voltage"
                    currentGraphName = "Feeder4Lateral2StartV"
                Case "Lateral 3 Start Voltage"
                    currentGraphName = "Feeder4Lateral3StartV"
                Case "Lateral 4 Start Voltage"
                    currentGraphName = "Feeder4Lateral4StartV"
                Case "Lateral 1 End Voltage"
                    currentGraphName = "Feeder4Lateral1EndV"
                Case "Lateral 2 End Voltage"
                    currentGraphName = "Feeder4Lateral2EndV"
                Case "Lateral 3 End Voltage"
                    currentGraphName = "Feeder4Lateral3EndV"
                Case "Lateral 4 End Voltage"
                    currentGraphName = "Feeder4Lateral4EndV"

                Case "Feeder 4 Start Current"
                    currentGraphName = "Feeder4StartI"
                    
                Case "Lateral 1 Start Current"
                    currentGraphName = "Feeder4Lateral1StartI"
                Case "Lateral 2 Start Current"
                    currentGraphName = "Feeder4Lateral2StartI"
                Case "Lateral 3 Start Current"
                    currentGraphName = "Feeder4Lateral3StartI"
                Case "Lateral 4 Start Current"
                    currentGraphName = "Feeder4Lateral4StartI"
            End Select
            
            ' Select the graph, copy it, change workingsheet
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartArea.Copy
            Set workingsheet = Sheets("Select Graphs")
            ' Determine the position of the new graph
            tempInt = j * 20
            yLocation = j * 20 + 19
            If tempInt = 0 Then tempInt = 1
            ' Paste it, and change the title
            workingsheet.Paste workingsheet.Range("D" & tempInt)
            Set graph = workingsheet.ChartObjects(currentGraphName)
            graph.Chart.ChartTitle.Text = "Feeder 4 " & graph.Chart.ChartTitle.Text
            
            Set RngToCover = workingsheet.Range("D" & tempInt & ":O" & yLocation)
            With graph
                .Height = RngToCover.Height
                .Width = RngToCover.Width
                .Top = RngToCover.Top
                .Left = RngToCover.Left
            End With
            j = j + 1
        End If
    Next i
    
    ' Change the active sheet to "Select Graphs", and hide the graphs form view
    Set workingsheet = Sheets("Select Graphs")
    workingsheet.Activate
    SelectGraphsForm.Hide
    
End Sub

Private Sub SelectGraphsFeederOneList_Click()

End Sub

Private Sub UserForm_Initialize()
    'For i = 1 To SelectGraphsForm.Controls.Count
        
    'Next i
    'Create ListBox
    For i = 1 To feeders
        Set NewListBox = settingsForm.designer.Controls.Add("Forms.listbox.1")
        With NewListBox
            .name = "fdr_1"
            .Top = 10
            .Left = i * 3 + (i - 1) * 105
            .Width = 105
            .Height = 180
            .Font.Size = 8
            .Font.name = "Tahoma"
            .BorderStyle = fmBorderStyleOpaque
            .SpecialEffect = fmSpecialEffectSunken
        End With
    Next i
    
    With SelectGraphsList
        .AddItem "Transformer Power Output"
        .AddItem "BusBar Voltage"
    End With
    
    With SelectGraphsFeederOneList
        .AddItem "Lateral 1 Start Voltage"
        .AddItem "Lateral 2 Start Voltage"
        .AddItem "Lateral 3 Start Voltage"
        .AddItem "Lateral 4 Start Voltage"
        
        .AddItem "Lateral 1 End Voltage"
        .AddItem "Lateral 2 End Voltage"
        .AddItem "Lateral 3 End Voltage"
        .AddItem "Lateral 4 End Voltage"
        
        .AddItem "Feeder 1 Start Current"
        
        .AddItem "Lateral 1 Start Current"
        .AddItem "Lateral 2 Start Current"
        .AddItem "Lateral 3 Start Current"
        .AddItem "Lateral 4 Start Current"
    End With
    
    With SelectGraphsFeederTwoList
        .AddItem "Lateral 1 Start Voltage"
        .AddItem "Lateral 2 Start Voltage"
        .AddItem "Lateral 3 Start Voltage"
        .AddItem "Lateral 4 Start Voltage"
        
        .AddItem "Lateral 1 End Voltage"
        .AddItem "Lateral 2 End Voltage"
        .AddItem "Lateral 3 End Voltage"
        .AddItem "Lateral 4 End Voltage"
        
        .AddItem "Feeder 2 Start Current"
        
        .AddItem "Lateral 1 Start Current"
        .AddItem "Lateral 2 Start Current"
        .AddItem "Lateral 3 Start Current"
        .AddItem "Lateral 4 Start Current"
    End With
    
    With SelectGraphsFeederThreeList
        .AddItem "Lateral 1 Start Voltage"
        .AddItem "Lateral 2 Start Voltage"
        .AddItem "Lateral 3 Start Voltage"
        .AddItem "Lateral 4 Start Voltage"
        
        .AddItem "Lateral 1 End Voltage"
        .AddItem "Lateral 2 End Voltage"
        .AddItem "Lateral 3 End Voltage"
        .AddItem "Lateral 4 End Voltage"
        
        .AddItem "Feeder 3 Start Current"
        
        .AddItem "Lateral 1 Start Current"
        .AddItem "Lateral 2 Start Current"
        .AddItem "Lateral 3 Start Current"
        .AddItem "Lateral 4 Start Current"
    End With
    
    With SelectGraphsFeederFourList
        .AddItem "Lateral 1 Start Voltage"
        .AddItem "Lateral 2 Start Voltage"
        .AddItem "Lateral 3 Start Voltage"
        .AddItem "Lateral 4 Start Voltage"
        
        .AddItem "Lateral 1 End Voltage"
        .AddItem "Lateral 2 End Voltage"
        .AddItem "Lateral 3 End Voltage"
        .AddItem "Lateral 4 End Voltage"
        
        .AddItem "Feeder 4 Start Current"
        
        .AddItem "Lateral 1 Start Current"
        .AddItem "Lateral 2 Start Current"
        .AddItem "Lateral 3 Start Current"
        .AddItem "Lateral 4 Start Current"
    End With
End Sub
