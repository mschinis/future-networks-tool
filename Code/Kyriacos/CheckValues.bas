Attribute VB_Name = "CheckValues"
Public NotCompliant As Integer

Public MaxTransformerUse As Double
Public MinTransformerUse As Double
Public MaxCurrentUseFeeder As Double
Public MinCurrentUseFeeder As Double
Public MaxCurrentUseLateral As Double
Public MinCurrentUseLateral As Double
Public MaxVoltage As Double
Public MinVoltage As Double
Public VoltageCompliance As Double

Public Sub CheckValuesPreset(ByVal NoCustomers As Integer, ByVal iter As Integer, ByRef TransformerArray() As Double, ByRef Feeders() As Double, ByRef Laterals() As Double, ByRef CustomersVoltages() As Double, ByRef CustomersLimits() As Byte)
    
    Dim TempArray As Variant
    Dim dValue As Double
    Network = PresetNetwork.Network
    
    If Network = "Urban" Then
        TransformerMax = 800
        If ChooseNetwork.TdayVal.Value <= 4 Or ChooseNetwork.TdayVal.Value >= 11 Then
            FeederCurrentMax = 309
            LateralCurrentMax = 209
        Else
            FeederCurrentMax = 297
            LateralCurrentMax = 202
        End If
        
    ElseIf Network = "SemiUrban" Then
        TransformerMax = 500
        If ChooseNetwork.TdayVal.Value <= 4 Or ChooseNetwork.TdayVal.Value >= 11 Then
            FeederCurrentMax = 309
            LateralCurrentMax = 209
        Else
            FeederCurrentMax = 297
            LateralCurrentMax = 202
        End If
        
    ElseIf Network = "Rural" Then
        TransformerMax = 200
        If ChooseNetwork.TdayVal.Value <= 4 Or ChooseNetwork.TdayVal.Value >= 11 Then
            FeederCurrentMax = 404
            LateralCurrentMax = 263
        Else
            FeederCurrentMax = 350
            LateralCurrentMax = 230
        End If
    End If
        
        
        'Check Transformer
        DSSCircuit.SetActiveElement ("transformer.LV_Transformer")
        TempArray = DSSCircuit.ActiveCktElement.Powers
        A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5
        B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5
        C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5
        TransformerUse = (A + B + C)

        TransformerArray(iter, 1) = TransformerUse
        
        TransformerUse = TransformerUse / TransformerMax
        If TransformerUse > 1 Then
        End If
        
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

        For i = 1 To 4 'Feeder Number
            
            'Check Currents at Start of the Feeder
            DSSCircuit.SetActiveElement ("Line.Feeder" & i & ".1")
            TempArray = DSSCircuit.ActiveCktElement.Currents
            A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5
            B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5
            C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5

            Feeders(iter, i, 1) = A
            Feeders(iter, i, 2) = B
            Feeders(iter, i, 3) = C
                    
            A = A / FeederCurrentMax
            B = B / FeederCurrentMax
            C = C / FeederCurrentMax
            
            If A > MaxCurrentUseFeeder Then MaxCurrentUseFeeder = A
            If B > MaxCurrentUseFeeder Then MaxCurrentUseFeeder = B
            If C > MaxCurrentUseFeeder Then MaxCurrentUseFeeder = C
            If A < MinCurrentUseFeeder Then MinCurrentUseFeeder = A
            If B < MinCurrentUseFeeder Then MinCurrentUseFeeder = B
            If C < MinCurrentUseFeeder Then MinCurrentUseFeeder = C
            
            If A > 1 Then
            End If
            If B > 1 Then
            End If
            If C > 1 Then
            End If
            
            For y = 1 To 4 'Lateral Number
                
                'Check Currents at Start of Lateral
                DSSCircuit.SetActiveElement ("Line.Lateral" & i & "_start_" & y)
                TempArray = DSSCircuit.ActiveCktElement.Currents
                A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5
                B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5
                C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5
    
                Laterals(iter, i, y, 1) = A
                Laterals(iter, i, y, 2) = B
                Laterals(iter, i, y, 3) = C
                    
                A = A / LateralCurrentMax
                B = B / LateralCurrentMax
                C = C / LateralCurrentMax
                
                If A > MaxCurrentUseLateral Then MaxCurrentUseLateral = A
                If B > MaxCurrentUseLateral Then MaxCurrentUseLateral = B
                If C > MaxCurrentUseLateral Then MaxCurrentUseLateral = C
                If A < MinCurrentUseLateral Then MinCurrentUseLateral = A
                If B < MinCurrentUseLateral Then MinCurrentUseLateral = B
                If C < MinCurrentUseLateral Then MinCurrentUseLateral = C
                
                If A > 1 Then
                End If
                If B > 1 Then
                End If
                If C > 1 Then
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
            
            For Z = 1 To (NoCustomers / 4)
                DSSCircuit.SetActiveElement ("Line.Consumer" & i & "_" & Z)
                TempArray = DSSCircuit.ActiveCktElement.Voltages
                A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5 / 230
                
                CustomersVoltages(i, Z, iter) = A
                
                dValue = 0
                If A > 1.1 Or A < 0.9 Then
                    CustomersLimits(i, Z, iter) = 1
                    NotCompliant = NotCompliant + 1
                ElseIf iter > 10 Then
                    For j = 1 To 10
                        dValue = CustomersVoltages(i, Z, iter - j) + dValue
                    Next
                    dValue = dValue / 10
                    If dValue < 0.94 Then
                        CustomersLimits(i, Z, iter) = 1
                        NotCompliant = NotCompliant + 1
                    End If
                End If
                

                
            Next
        Next
    

End Sub

Public Sub Check_Compliance()

    Dim compliance As Double
    Dim maxcompliant As Long
    
    
    maxcompliant = CLng(PresetNetwork.customers) * CLng(Start.RunHours)
    VoltageCompliance = (maxcompliant - CheckValues.NotCompliant) / maxcompliant
    

End Sub
