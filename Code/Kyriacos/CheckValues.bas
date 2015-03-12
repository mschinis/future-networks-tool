Attribute VB_Name = "CheckValues"
Public compliance As Integer

Public MaxTransformerUse As Double
Public MinTransformerUse As Double
Public MaxCurrentUseFeeder As Double
Public MinCurrentUseFeeder As Double
Public MaxCurrentUseLateral As Double
Public MinCurrentUseLateral As Double
Public MaxVoltage As Double
Public MinVoltage As Double
Public VoltageCompliance As Double
Public PercentageCustomersVoltage As Double

Public TransformerMax As Integer
Public feedercurrentmax As Integer
Public lateralcurrentmax As Integer
Public VoltageMax As Double
Public VoltageMin As Double
Public VoltageAverageMin As Double



Public Sub CheckValuesPreset(ByVal NoCustomers As Integer, ByVal iter As Integer, ByRef TransformerArray() As Double, ByRef Feeders() As Double, ByRef Laterals() As Double, ByRef CustomersVoltages() As Double, ByRef CustomersLimits() As Byte, ByRef CurrentFlags() As Byte)
    
    Dim TempArray As Variant
    Dim dValue As Double
    Dim TransformerOverloaded As Boolean
    Dim A, B, C As Double

    
    Network = PresetNetwork.Network
    
'    If Start.OverrideDefault = False Then
        VoltageMax = AdvancedProperties.VoltageMax
        VoltageMin = AdvancedProperties.VoltageMin
        VoltageAverageMin = AdvancedProperties.VoltageAverageMin
    
        If Network = "Urban" Then
            TransformerMax = 800 * AdvancedProperties.TransformerMax / 100

            If ChooseNetwork.MonthVal.Value <= 4 Or ChooseNetwork.MonthVal.Value >= 11 Then
                feedercurrentmax = 309 * AdvancedProperties.FeederMax / 100
                lateralcurrentmax = 209 * AdvancedProperties.LateralMax / 100
            Else
                feedercurrentmax = 297 * AdvancedProperties.FeederMax / 100
                lateralcurrentmax = 202 * AdvancedProperties.LateralMax / 100
            End If
        
        ElseIf Network = "SemiUrban" Then
            TransformerMax = 500 * AdvancedProperties.TransformerMax / 100
            If ChooseNetwork.MonthVal.Value <= 4 Or ChooseNetwork.MonthVal.Value >= 11 Then
                feedercurrentmax = 309 * AdvancedProperties.FeederMax / 100
                lateralcurrentmax = 209 * AdvancedProperties.LateralMax / 100
            Else
                feedercurrentmax = 297 * AdvancedProperties.FeederMax / 100
                lateralcurrentmax = 202 * AdvancedProperties.LateralMax / 100
            End If
        
        ElseIf Network = "Rural" Then
            TransformerMax = 200 * AdvancedProperties.TransformerMax / 100
            If ChooseNetwork.MonthVal.Value <= 4 Or ChooseNetwork.MonthVal.Value >= 11 Then
                feedercurrentmax = 404 * AdvancedProperties.FeederMax / 100
                lateralcurrentmax = 263 * AdvancedProperties.LateralMax / 100
            Else
                feedercurrentmax = 350 * AdvancedProperties.FeederMax / 100
                lateralcurrentmax = 230 * AdvancedProperties.LateralMax / 100
            End If
        End If
    
'    Else
'
'        feedercurrentmax = AdvancedProperties.FeederMax
'        lateralcurrentmax = AdvancedProperties.LateralMax
'        TransformerMax = AdvancedProperties.TransformerMax
'        VoltageMax = AdvancedProperties.VoltageMax
'        VoltageMin = AdvancedProperties.VoltageMin
'        VoltageAverageMin = AdvancedProperties.VoltageAverageMin
'
'    End If
        
        
        'Check Transformer
        DSSCircuit.SetActiveElement ("transformer.LV_Transformer")
        TempArray = DSSCircuit.ActiveCktElement.Powers
        A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5
        B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5
        C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5
        TransformerUse = (A + B + C)

        TransformerArray(iter, 1) = TransformerUse
        If (TempArray(LBound(TempArray))) + (TempArray(LBound(TempArray) + 2)) + (TempArray(LBound(TempArray) + 4)) < 0 Then TransformerArray(iter, 1) = -TransformerUse
        
        TransformerUse = TransformerUse / TransformerMax
        
        If TransformerUse > MaxTransformerUse Then MaxTransformerUse = TransformerUse
        If TransformerUse < MinTransformerUse Then MinTransformerUse = TransformerUse
    

        'Check Voltages on Busbar
        DSSCircuit.SetActiveElement ("Line.Feeder1.1")
        TempArray = DSSCircuit.ActiveCktElement.Voltages
        A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5 / 230
        B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5 / 230
        C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5 / 230
        
        TransformerArray(iter, 2) = A
        TransformerArray(iter, 3) = B
        TransformerArray(iter, 4) = C
        
        If A > MaxVoltage Then MaxVoltage = A
        If B > MaxVoltage Then MaxVoltage = B
        If C > MaxVoltage Then MaxVoltage = C
        If A < MinVoltage Then MinVoltage = A
        If B < MinVoltage Then MinVoltage = B
        If C < MinVoltage Then MinVoltage = C
        
        If A > 1.1 Or A < 0.94 Then
        End If
        If B > 1.1 Or B < 0.94 Then
        End If
        If C > 1.1 Or C < 0.94 Then
        End If

        For i = 1 To Assign_Profiles.NoFeeders 'Feeder Number
            
            'Check Currents at Start of the Feeder
            DSSCircuit.SetActiveElement ("Line.Feeder" & i & ".1")
            TempArray = DSSCircuit.ActiveCktElement.Currents
            A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5
            B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5
            C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5

            Feeders(iter, i, 1) = A
            Feeders(iter, i, 2) = B
            Feeders(iter, i, 3) = C
                    
            A = A / feedercurrentmax
            B = B / feedercurrentmax
            C = C / feedercurrentmax
            
            If A > MaxCurrentUseFeeder Then MaxCurrentUseFeeder = A
            If B > MaxCurrentUseFeeder Then MaxCurrentUseFeeder = B
            If C > MaxCurrentUseFeeder Then MaxCurrentUseFeeder = C
            If A < MinCurrentUseFeeder Then MinCurrentUseFeeder = A
            If B < MinCurrentUseFeeder Then MinCurrentUseFeeder = B
            If C < MinCurrentUseFeeder Then MinCurrentUseFeeder = C
            
            If A > 1 Then
                CurrentFlags(i, 1) = 1
            End If
            If B > 1 Then
                CurrentFlags(i, 1) = 1
            End If
            If C > 1 Then
                CurrentFlags(i, 1) = 1
            End If
            
            For y = 1 To Assign_Profiles.NoLaterals 'Lateral Number
                
                'Check Currents at Start of Lateral
                DSSCircuit.SetActiveElement ("Line.Lateral" & i & "_start_" & y)
                TempArray = DSSCircuit.ActiveCktElement.Currents
                A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5
                B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5
                C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5
    
                Laterals(iter, i, y, 1) = A
                Laterals(iter, i, y, 2) = B
                Laterals(iter, i, y, 3) = C
                    
                A = A / lateralcurrentmax
                B = B / lateralcurrentmax
                C = C / lateralcurrentmax
                
                If A > MaxCurrentUseLateral Then MaxCurrentUseLateral = A
                If B > MaxCurrentUseLateral Then MaxCurrentUseLateral = B
                If C > MaxCurrentUseLateral Then MaxCurrentUseLateral = C
                If A < MinCurrentUseLateral Then MinCurrentUseLateral = A
                If B < MinCurrentUseLateral Then MinCurrentUseLateral = B
                If C < MinCurrentUseLateral Then MinCurrentUseLateral = C
                
                If A > 1 Then
                    CurrentFlags(i, y + 1) = 1
                End If
                If B > 1 Then
                    CurrentFlags(i, y + 1) = 1
                End If
                If C > 1 Then
                    CurrentFlags(i, y + 1) = 1
                End If
                
                'Check Voltages at Start of Lateral
                DSSCircuit.SetActiveElement ("Line.Lateral" & i & "_start_" & y)
                TempArray = DSSCircuit.ActiveCktElement.Voltages
                A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5 / 230
                B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5 / 230
                C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5 / 230
    
                Laterals(iter, i, y, 4) = A
                Laterals(iter, i, y, 5) = B
                Laterals(iter, i, y, 6) = C
                
                If A > MaxVoltage Then MaxVoltage = A
                If B > MaxVoltage Then MaxVoltage = B
                If C > MaxVoltage Then MaxVoltage = C
                If A < MinVoltage Then MinVoltage = A
                If B < MinVoltage Then MinVoltage = B
                If C < MinVoltage Then MinVoltage = C
                    
                If A > 1.1 Or A < 0.94 Then
                End If
                If B > 1.1 Or B < 0.94 Then
                End If
                If C > 1.1 Or C < 0.94 Then
                End If

                'Check Voltages at End of Lateral
                DSSCircuit.SetActiveElement ("Line.Lateral" & i & "_end_" & y)
                TempArray = DSSCircuit.ActiveCktElement.Voltages
                A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5 / 230
                B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5 / 230
                C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5 / 230
    
                Laterals(iter, i, y, 7) = A
                Laterals(iter, i, y, 8) = B
                Laterals(iter, i, y, 9) = C
                
                If A > MaxVoltage Then MaxVoltage = A
                If B > MaxVoltage Then MaxVoltage = B
                If C > MaxVoltage Then MaxVoltage = C
                If A < MinVoltage Then MinVoltage = A
                If B < MinVoltage Then MinVoltage = B
                If C < MinVoltage Then MinVoltage = C
                    
                If A > 1.1 Or A < 0.94 Then
                End If
                If B > 1.1 Or B < 0.94 Then
                End If
                If C > 1.1 Or C < 0.94 Then
                End If

          
            Next
            
            For z = 1 To (NoCustomers / 4)
                DSSCircuit.SetActiveElement ("Line.Consumer" & i & "_" & z)
                TempArray = DSSCircuit.ActiveCktElement.Voltages
                A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5 / 230

                CustomersVoltages(i, z, iter) = A

                dValue = 0
                If A > VoltageMax Or A < VoltageMin Then
                    CustomersLimits(i, z, iter) = 1
                    Start.NotCompliant(z + ((i * NoCustomers / 4) - NoCustomers / 4)) = Start.NotCompliant(z + ((i * NoCustomers / 4) - NoCustomers / 4)) + 1
                    Start.CustomerVoltageLimit(z + ((i * NoCustomers / 4) - NoCustomers / 4)) = 1
                ElseIf iter > 10 Then
                    For j = 1 To 10
                        dValue = CustomersVoltages(i, z, iter - j) + dValue
                    Next
                    dValue = dValue / 10
                    If dValue < VoltageAverageMin Then
                        CustomersLimits(i, z, iter) = 1
                        Start.NotCompliant(z + ((i * NoCustomers / 4) - NoCustomers / 4)) = Start.NotCompliant(z + ((i * NoCustomers / 4) - NoCustomers / 4)) + 1
                        Start.CustomerVoltageLimit(z + ((i * NoCustomers / 4) - NoCustomers / 4)) = 1
                    End If
                End If



            Next
        Next
    

End Sub

Public Sub Check_Compliance()

    Dim maxcompliant As Long
    compliance = 0
    VoltageCompliance = 0
    
    For i = 1 To PresetNetwork.customers
        Start.NotCompliant(i) = Start.NotCompliant(i) / Start.RunHours
        If Start.NotCompliant(i) > 0.05 Then VoltageCompliance = VoltageCompliance + 1
    Next
    
    VoltageCompliance = (PresetNetwork.customers - VoltageCompliance) / PresetNetwork.customers
    
'    maxcompliant = CLng(PresetNetwork.customers) * CLng(Start.RunHours)
'    VoltageCompliance = (maxcompliant - CheckValues.NotCompliant) / maxcompliant
    

End Sub

Public Sub Customer_Voltage_Percentage()

    Dim adder As Integer
    adder = 0
    
    For i = 1 To PresetNetwork.customers
    
        adder = adder + Start.CustomerVoltageLimit(i)
    
    Next
    
    PercentageCustomersVoltage = adder / PresetNetwork.customers
    
End Sub
