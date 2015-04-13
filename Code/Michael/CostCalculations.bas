Attribute VB_Name = "CostCalculations"
Dim voltageCellsMax() As String
Dim voltageCellsMin() As String
Dim voltageCostSummary() As String

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
    CostCalculations.FeederCheck
    CostCalculations.LateralCheck
    CostCalculations.VoltageCheck
    
    
    End Sub

Public Sub TransformerCheck()
    Dim workingsheet As Worksheet
    Set workingsheet = Worksheets("Results Summary")
    TransformerUsage = workingsheet.Range("C13")
    
    Set workingsheet = Worksheets("Limits")
    transformerLimit = workingsheet.Range("G4")
    
    If TransformerUsage * transformerLimit < transformerLimit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C5").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C5").Value = "Yes"
    End If

End Sub

Public Sub FeederCheck()
    Dim workingsheet As Worksheet
    Set workingsheet = Worksheets("FeederCurrentRollingAverages")
    Feeder1current = workingsheet.Range("C1390")
    
    Set workingsheet = Worksheets("Limits")
    Feeder1Limit = workingsheet.Range("E4")
    
    If Feeder1current * Feeder1Limit < Feeder1Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C9").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C9").Value = "Yes"
    End If
    
    
    Set workingsheet = Worksheets("FeederCurrentRollingAverages")
    Feeder2current = workingsheet.Range("F1390")
    
    Set workingsheet = Worksheets("Limits")
    Feeder2Limit = workingsheet.Range("E4")
    
    If Feeder2current * Feeder2Limit < Feeder2Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C15").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C15").Value = "Yes"
    End If
    
    
    Set workingsheet = Worksheets("FeederCurrentRollingAverages")
    Feeder3current = workingsheet.Range("I1390")
    
    Set workingsheet = Worksheets("Limits")
    Feeder3Limit = workingsheet.Range("E4")
    
    If Feeder3current * Feeder3Limit < Feeder3Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C21").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C21").Value = "Yes"
    End If
    
    
    Set workingsheet = Worksheets("FeederCurrentRollingAverages")
    Feeder4current = workingsheet.Range("L1390")
    
    Set workingsheet = Worksheets("Limits")
    Feeder4Limit = workingsheet.Range("E4")
    
    If Feeder4current * Feeder4Limit < Feeder4Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C27").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C27").Value = "Yes"
    End If

End Sub

Public Sub LateralCheck()
    Dim workingsheet As Worksheet
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral1current = workingsheet.Range("C1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral1Limit = workingsheet.Range("D4")
    
    If Lateral1current * Lateral1Limit < Lateral1Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C10").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C10").Value = "Yes"
    End If
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral2current = workingsheet.Range("F1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral2Limit = workingsheet.Range("D4")
    
    If Lateral2current * Lateral2Limit < Lateral2Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C11").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C11").Value = "Yes"
    End If
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral3current = workingsheet.Range("I1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral3Limit = workingsheet.Range("D4")
    
    If Lateral3current * Lateral3Limit < Lateral3Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C12").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C12").Value = "Yes"
    End If
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral4current = workingsheet.Range("L1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral4Limit = workingsheet.Range("D4")
    
    If Lateral4current * Lateral4Limit < Lateral4Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C13").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C13").Value = "Yes"
    End If
    
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral5current = workingsheet.Range("O1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral5Limit = workingsheet.Range("D4")
    
    If Lateral5current * Lateral5Limit < Lateral5Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C16").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C16").Value = "Yes"
    End If
    
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral6current = workingsheet.Range("R1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral6Limit = workingsheet.Range("D4")
    
    If Lateral6current * Lateral6Limit < Lateral6Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C17").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C17").Value = "Yes"
    End If
    
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral7current = workingsheet.Range("U1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral7Limit = workingsheet.Range("D4")
    
    If Lateral7current * Lateral7Limit < Lateral7Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C18").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C18").Value = "Yes"
    End If
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral8current = workingsheet.Range("X1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral8Limit = workingsheet.Range("D4")
    
    If Lateral8current * Lateral8Limit < Lateral8Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C19").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C19").Value = "Yes"
    End If
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral9current = workingsheet.Range("AA1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral9Limit = workingsheet.Range("D4")
    
    If Lateral9current * Lateral9Limit < Lateral9Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C22").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C22").Value = "Yes"
    End If
    
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral10current = workingsheet.Range("AD1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral10Limit = workingsheet.Range("D4")
    
    If Lateral10current * Lateral10Limit < Lateral10Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C23").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C23").Value = "Yes"
    End If
    
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral11current = workingsheet.Range("AG1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral11Limit = workingsheet.Range("D4")
    
    If Lateral11current * Lateral11Limit < Lateral11Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C24").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C24").Value = "Yes"
    End If
    
    
    
     Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral12current = workingsheet.Range("AJ1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral12Limit = workingsheet.Range("D4")
    
    If Lateral12current * Lateral12Limit < Lateral12Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C25").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C25").Value = "Yes"
    End If
    
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral13current = workingsheet.Range("AM1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral13Limit = workingsheet.Range("D4")
    
    If Lateral13current * Lateral13Limit < Lateral13Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C28").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C28").Value = "Yes"
    End If
    
    
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral14current = workingsheet.Range("AP1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral14Limit = workingsheet.Range("D4")
    
    If Lateral14current * Lateral14Limit < Lateral14Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C29").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C29").Value = "Yes"
    End If
    
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral15current = workingsheet.Range("AS1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral15Limit = workingsheet.Range("D4")
    
    If Lateral15current * Lateral15Limit < Lateral15Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C30").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C30").Value = "Yes"
    End If
    
    
    
    Set workingsheet = Worksheets("CurrentRollingAverages")
    Lateral16current = workingsheet.Range("AV1392")
    
    Set workingsheet = Worksheets("Limits")
    Lateral16Limit = workingsheet.Range("D4")
    
    If Lateral16current * Lateral16Limit < Lateral16Limit Then
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C31").Value = "No"
    Else
        Set workingsheet = Worksheets("Cost Summary")
        workingsheet.Range("C31").Value = "Yes"
    End If
End Sub
