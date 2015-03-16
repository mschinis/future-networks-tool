Attribute VB_Name = "CostCalculations"
Dim voltageCellsMax() As String
Dim voltageCellsMin() As String
Dim voltageCostSummary() As String
Dim FeederCells() As String
Dim FeederCostSummary() As String
Dim LateralCells() As String
Dim LateralCostSummary() As String



Public Sub VoltageCellAllocation()
    ReDim voltageCostSummary(1 To 33)
    ReDim voltageCellsMax(1 To 33)
    ReDim voltageCellsMin(1 To 33)
    
    voltageCellsMax(1) = "C1442"
    voltageCellsMax(2) = "F1442"
    voltageCellsMax(3) = "I1442"
    voltageCellsMax(4) = "L1442"
    voltageCellsMax(5) = "O1442"
    voltageCellsMax(6) = "R1442"
    voltageCellsMax(7) = "U1442"
    voltageCellsMax(8) = "X1442"
    voltageCellsMax(9) = "AA1442"
    voltageCellsMax(10) = "AD1442"
    voltageCellsMax(11) = "AG1442"
    voltageCellsMax(12) = "AJ1442"
    voltageCellsMax(13) = "AM1442"
    voltageCellsMax(14) = "AP1442"
    voltageCellsMax(15) = "AS1442"
    voltageCellsMax(16) = "AV1442"
    voltageCellsMax(17) = "AY1442"
    voltageCellsMax(18) = "BB1442"
    voltageCellsMax(19) = "BE1442"
    voltageCellsMax(20) = "BH1442"
    voltageCellsMax(21) = "BK1442"
    voltageCellsMax(22) = "BN1442"
    voltageCellsMax(23) = "BQ1442"
    voltageCellsMax(24) = "BT1442"
    voltageCellsMax(25) = "BW1442"
    voltageCellsMax(26) = "BZ1442"
    voltageCellsMax(27) = "CC1442"
    voltageCellsMax(28) = "CF1442"
    voltageCellsMax(29) = "CI1442"
    voltageCellsMax(30) = "CL1442"
    voltageCellsMax(31) = "CO1442"
    voltageCellsMax(32) = "CR1442"
    voltageCellsMax(33) = "CU1442"
    
    voltageCellsMin(1) = "C1441"
    voltageCellsMin(2) = "F1441"
    voltageCellsMin(3) = "I1441"
    voltageCellsMin(4) = "L1441"
    voltageCellsMin(5) = "O1441"
    voltageCellsMin(6) = "R1441"
    voltageCellsMin(7) = "U1441"
    voltageCellsMin(8) = "X1441"
    voltageCellsMin(9) = "AA1441"
    voltageCellsMin(10) = "AD1441"
    voltageCellsMin(11) = "AG1441"
    voltageCellsMin(12) = "AJ1441"
    voltageCellsMin(13) = "AM1441"
    voltageCellsMin(14) = "AP1441"
    voltageCellsMin(15) = "AS1441"
    voltageCellsMin(16) = "AV1441"
    voltageCellsMin(17) = "AY1441"
    voltageCellsMin(18) = "BB1441"
    voltageCellsMin(19) = "BE1441"
    voltageCellsMin(20) = "BH1441"
    voltageCellsMin(21) = "BK1441"
    voltageCellsMin(22) = "BN1441"
    voltageCellsMin(23) = "BQ1441"
    voltageCellsMin(24) = "BT1441"
    voltageCellsMin(25) = "BW1441"
    voltageCellsMin(26) = "BZ1441"
    voltageCellsMin(27) = "CC1441"
    voltageCellsMin(28) = "CF1441"
    voltageCellsMin(29) = "CI1441"
    voltageCellsMin(30) = "CL1441"
    voltageCellsMin(31) = "CO1441"
    voltageCellsMin(32) = "CR1441"
    voltageCellsMin(33) = "CU1441"
    
    voltageCostSummary(1) = "C35"
    voltageCostSummary(2) = "C38"
    voltageCostSummary(3) = "C39"
    voltageCostSummary(4) = "C40"
    voltageCostSummary(5) = "C41"
    voltageCostSummary(6) = "C43"
    voltageCostSummary(7) = "C44"
    voltageCostSummary(8) = "C45"
    voltageCostSummary(9) = "C46"
    voltageCostSummary(10) = "C48"
    voltageCostSummary(11) = "C49"
    voltageCostSummary(12) = "C50"
    voltageCostSummary(13) = "C51"
    voltageCostSummary(14) = "C53"
    voltageCostSummary(15) = "C54"
    voltageCostSummary(16) = "C55"
    voltageCostSummary(17) = "C56"
    voltageCostSummary(18) = "C59"
    voltageCostSummary(19) = "C60"
    voltageCostSummary(20) = "C61"
    voltageCostSummary(21) = "C62"
    voltageCostSummary(22) = "C64"
    voltageCostSummary(23) = "C65"
    voltageCostSummary(24) = "C66"
    voltageCostSummary(25) = "C67"
    voltageCostSummary(26) = "C69"
    voltageCostSummary(27) = "C70"
    voltageCostSummary(28) = "C71"
    voltageCostSummary(29) = "C72"
    voltageCostSummary(30) = "C74"
    voltageCostSummary(31) = "C75"
    voltageCostSummary(32) = "C76"
    voltageCostSummary(33) = "C77"
    
End Sub

Public Sub VoltageCheck()

For i = 1 To 33
    Set workingsheet = Worksheets("VoltageRollingAverages")
    VoltageValueMin = workingsheet.Range(voltageCellsMin(i))
    VoltageValueMax = workingsheet.Range(voltageCellsMax(i))
    
    Set workingsheet = Worksheets("Limits")
    VoltageMinLimit = workingsheet.Range("C4")
    VoltageMaxLimit = workingsheet.Range("B4")
    
    If VoltageValueMin < VoltageMinLimit Or VoltageValueMax > VoltageMaxLimit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range(voltageCostSummary(i)).Value = "Yes"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range(voltageCostSummary(i)).Value = "No"
    End If
Next i

End Sub

Public Sub CalculateCosts()
    CostCalculations.VoltageCellAllocation
    CostCalculations.TransformerCheck
    CostCalculations.FeederCellAllocation
    CostCalculations.FeederCheck
    CostCalculations.LateralCellAllocation
    CostCalculations.LateralCheck
    CostCalculations.VoltageCheck
    
       
    
    End Sub

Public Sub TransformerCheck()
    Dim workingsheet As Worksheet
    Set workingsheet = Worksheets("Limits")
    Dim month As Integer
    Dim limit1, limit2, limit3 As Double
    Dim limitreached As Integer
    
    
    month = ChooseNetwork.MonthVal
    
    If month = 1 Or month = 2 Or month = 11 Or month = 12 Then
        limit1 = 1.52
        limit2 = 1.36
        limit3 = 1.3
    ElseIf month = 5 Or month = 9 Or month = 10 Or month = 4 Or month = 3 Then
        limit1 = 1.37
        limit2 = 1.24
        limit3 = 1.19
    Else
        limit1 = 1.18
        limit2 = 1.12
        limit3 = 1.1
    End If
    
    
    
    transformerlimit = workingsheet.Range("G4")
    
    Set workingsheet = Worksheets("PowerRollingAverages")
    transformerusage = workingsheet.Range("C1385")
    
    overloadHours = 0
    
    
    
  If transformerusage > transformerlimit Then
  
     If transformerusage < (limit3 * transformerlimit) Then
     
        For i = 2 To 1382
            If workingsheet.Range("C" & i).Value > transformerlimit Then
                overloadHours = overloadHours + 1


            End If
        Next i
            If overloadHours < 360 Then
                Set workingsheet = Worksheets("Cost Summary")
                workingsheet.Range("C5").Value = "No"
            Else
                Set workingsheet = Worksheets("Cost Summary")
                workingsheet.Range("C5").Value = "Yes"

            End If
     'count overload hours, if less than 6 hours, then ok, if not, then replace asset
        
        limitreached = 1
    ElseIf transformerusage < limit2 * transformerlimit Then
        For i = 2 To 1382
            If workingsheet.Range("C" & i).Value > transformerlimit Then
                overloadHours = overloadHours + 1


            End If
        Next i
            If overloadHours < 240 Then
                Set workingsheet = Worksheets("Cost Summary")
                workingsheet.Range("C5").Value = "No"
            Else
                Set workingsheet = Worksheets("Cost Summary")
                workingsheet.Range("C5").Value = "Yes"

            End If
    'countoverload hours, if less than 4 hours then ok, if not then replace asset.
    
    
        
        limitreached = 2
    ElseIf transformerusage < limit1 * transformerlimit Then
        For i = 2 To 1382
            If workingsheet.Range("C" & i).Value > transformerlimit Then
                overloadHours = overloadHours + 1


            End If
        Next i
            If overloadHours < 120 Then
                Set workingsheet = Worksheets("Cost Summary")
                workingsheet.Range("C5").Value = "No"
            Else
                Set workingsheet = Worksheets("Cost Summary")
                workingsheet.Range("C5").Value = "Yes"

            End If
  'count overload hours, if less than 2 hours then ok, if not then replace asset.
       
        
        limitreached = 3
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C5").Value = "Yes"
        'overloaded.
    End If
        
  Else
  
        limitreached = 0
    
  End If
 

End Sub
Public Sub FeederCellAllocation()
    ReDim FeederCells(1 To 4)
    ReDim FeederCostSummary(1 To 4)
    
    FeederCells(1) = "C1390"
    FeederCells(2) = "F1390"
    FeederCells(3) = "I1390"
    FeederCells(4) = "L1390"
    
    FeederCostSummary(1) = "C9"
    FeederCostSummary(2) = "C15"
    FeederCostSummary(3) = "C21"
    FeederCostSummary(4) = "C27"
     
End Sub


Public Sub FeederCheck()

For i = 1 To 4
    Set workingsheet = Worksheets("FeederCurrentRollingAverages")
    FeederValue = workingsheet.Range(FeederCells(i))
    
    Set workingsheet = Worksheets("Limits")
    FeederLimit = workingsheet.Range("E4")
    
    If FeederValue * FeederLimit > FeederLimit * 1.15 Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range(FeederCostSummary(i)).Value = "Yes"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range(FeederCostSummary(i)).Value = "No"
    End If
Next i
     
End Sub


Public Sub LateralCellAllocation()

    ReDim LateralCells(1 To 16)
    ReDim LateralCostSummary(1 To 16)
    
    LateralCells(1) = "C1392"
    LateralCells(2) = "F1392"
    LateralCells(3) = "I1392"
    LateralCells(4) = "L1392"
    LateralCells(5) = "O1392"
    LateralCells(6) = "R1392"
    LateralCells(7) = "U1392"
    LateralCells(8) = "X1392"
    LateralCells(9) = "AA1392"
    LateralCells(10) = "AD1392"
    LateralCells(11) = "AG1392"
    LateralCells(12) = "AJ1392"
    LateralCells(13) = "AM1392"
    LateralCells(14) = "AP1392"
    LateralCells(15) = "AS1392"
    LateralCells(16) = "AV1392"
    
    LateralCostSummary(1) = "C10"
    LateralCostSummary(2) = "C11"
    LateralCostSummary(3) = "C12"
    LateralCostSummary(4) = "C13"
    LateralCostSummary(5) = "C16"
    LateralCostSummary(6) = "C17"
    LateralCostSummary(7) = "C18"
    LateralCostSummary(8) = "C19"
    LateralCostSummary(9) = "C22"
    LateralCostSummary(10) = "C23"
    LateralCostSummary(11) = "C24"
    LateralCostSummary(12) = "C25"
    LateralCostSummary(13) = "C28"
    LateralCostSummary(14) = "C29"
    LateralCostSummary(15) = "C30"
    LateralCostSummary(16) = "C31"
    

End Sub



Public Sub LateralCheck()

For i = 1 To 16
    Set workingsheet = Worksheets("CurrentRollingAverages")
    LateralValue = workingsheet.Range(LateralCells(i))
    
    Set workingsheet = Worksheets("Limits")
    LateralLimit = workingsheet.Range("D4")
    
    If LateralValue * LateralLimit > LateralLimit * 1.15 Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range(LateralCostSummary(i)).Value = "Yes"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range(LateralCostSummary(i)).Value = "No"
    End If
Next i
    
End Sub

