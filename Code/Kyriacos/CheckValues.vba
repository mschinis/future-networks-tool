Attribute VB_Name = "CheckValues"
   
Public Sub CheckValuesPreset(ByVal iter As Integer, ByRef TransformerArray() As Double, ByRef Feeders() As Double, ByRef Laterals() As Double)
    
    Dim TempArray As Variant
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
        TransformerUse = (kVAphaseA + kVAphaseB + kVAphaseC)

        TransformerArray(iter, 1) = TransformerUse
        
        TransformerUse = TransformerUse / TransformerMax
        If TransformerUse > 1 Then
        End If

        'Check Voltages on Busbar
        DSSCircuit.SetActiveElement ("Line.Feeder1.1")
        TempArray = DSSCircuit.ActiveCktElement.Voltages
        A = (TempArray(LBound(TempArray)) ^ 2 + TempArray(LBound(TempArray) + 1) ^ 2) ^ 0.5 / 230
        B = (TempArray(LBound(TempArray) + 2) ^ 2 + TempArray(LBound(TempArray) + 3) ^ 2) ^ 0.5 / 230
        C = (TempArray(LBound(TempArray) + 4) ^ 2 + TempArray(LBound(TempArray) + 5) ^ 2) ^ 0.5 / 230
        
        TransformerArray(iter, 2) = A
        TransformerArray(iter, 3) = B
        TransformerArray(iter, 4) = C
        
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
                    
                    
                If A > 1.1 Or A < 0.94 Then
                End If
                If B > 1.1 Or B < 0.94 Then
                End If
                If C > 1.1 Or C < 0.94 Then
                End If

          
            Next
        Next


End Sub
