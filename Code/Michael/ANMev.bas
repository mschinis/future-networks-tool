Attribute VB_Name = "ANMev"
Public MaxCharge() As Integer
Public Charge() As Integer
Public EVFlags() As Integer
Public achieved As Integer
Public achievedlaterals() As Integer
Public achievedfeeders() As Integer

Public MaxLaterals() As Integer
Public LateralsAssigned() As Integer
Public MaxFeeders() As Integer
Public FeedersAssigned() As Integer





Public Sub EVManagement(ByVal i As Integer)
    
    Call CheckEV
    Call CalculateDisconnectionsLaterals(i)
    
    If ChooseNetwork.EVANM = True Then
        
        Call LateralManagementEV(i, Start.Laterals)
        Call FeederManagementEV(i, Start.Feeders)
        Call TransformerManagementEV(i, Start.Feeders, Start.TransformerArray(i, 1) / CheckValues.TransformerMax)
    
    End If

    
End Sub

Public Sub LateralManagementEV(ByVal iter As Integer, ByRef CurrentUse() As Double)

Dim i, y, z, h As Integer
Dim lateralrequired As Integer

ReDim achievedlaterals(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 3)
ReDim achievedfeeders(1 To Assign_Profiles.NoLaterals, 1 To 3)
achieved = 0

For i = 1 To Assign_Profiles.NoFeeders
    For y = 1 To Assign_Profiles.NoLaterals
        For z = 1 To 3
            If CurrentUse(iter, i, y, z) / CheckValues.lateralcurrentmax > 1 Then
                For h = 1 To Assign_Profiles.NoEV
                    max = 0
                    lateralrequired = (CurrentUse(iter, i, y, z) - CheckValues.lateralcurrentmax) * 0.5 / 16
                    For A = 1 To Assign_Profiles.NoEV
                        If Assign_Profiles.EVLocation(1, A) = i And Assign_Profiles.EVLocation(2, A) = y And Assign_Profiles.EVLocation(3, A) = z Then
                            If EVFlags(A) = 1 Then
                                If max < Charge(A) Then
                                    max = Charge(A)
                                    comp = A
                                End If
                            End If
                        End If
                    Next
                
                    If achievedlaterals(i, y, z) < lateralrequired Then
                        If comp > 0 Then
                            If EVFlags(comp) = 1 Then
                                EVFlags(comp) = 4
                                achieved = achieved + 1
                                achievedlaterals(i, y, z) = achievedlaterals(i, y, z) + 1
                                achievedfeeders(i, z) = achievedfeeders(i, z) + 1
                                Call DisconnectEV(comp)
                            End If
                        End If
                    End If
                    If achievedlaterals(i, y, z) = lateralrequired Then
                        Exit For
                    End If
                Next
            End If
        Next
    Next
Next
                                
End Sub

Public Sub FeederManagementEV(ByVal iter As Integer, ByRef CurrentUse() As Double)

Dim i, y, z, h As Integer
Dim feederrequired As Integer

'ReDim achievedlaterals(1 To 4, 1 To 4, 1 To 3)
'ReDim achievedfeeder(1 To 4, 1 To 3)
'achieved = 0

For i = 1 To Assign_Profiles.NoFeeders
        For z = 1 To 3
            If CurrentUse(iter, i, z) / CheckValues.feedercurrentmax > 0.95 Then
                For h = 1 To Assign_Profiles.NoEV
                    max = 0
                    feederrequired = (CurrentUse(iter, i, z) - CheckValues.feedercurrentmax * 0.95) * 0.5 / 16
                    For A = 1 To Assign_Profiles.NoEV
                        If Assign_Profiles.EVLocation(1, A) = i And Assign_Profiles.EVLocation(3, A) = z Then
                            If EVFlags(A) = 1 Then
                                If max < Charge(A) Then
                                    max = Charge(A)
                                    comp = A
                                End If
                            End If
                        End If
                    Next
                
                    If achievedfeeders(i, z) < feederrequired Then
                        If comp > 0 Then
                            If EVFlags(comp) = 1 Then
                                EVFlags(comp) = 4
                                achieved = achieved + 1
                                achievedfeeders(i, z) = achievedfeeders(i, z) + 1
                                Call DisconnectEV(comp)
                            End If
                        End If
                    End If
                    If achievedfeeders(i, z) = feederrequired Then
                        Exit For
                    End If
                Next
            End If
        Next
Next



End Sub

Public Sub TransformerManagementEV(ByVal iter As Integer, ByRef feedercurrents() As Double, ByVal TransformerUse As Double)

Dim required As Integer
Dim max, min, comp, j, y, i, m, k As Integer
Dim upperlimit, lowerlimit As Double
Dim achievedcon As Integer

upperlimit = 0.96
lowerlimit = 0.93
achievedcon = 0

required = (((TransformerUse * CheckValues.TransformerMax) - CheckValues.TransformerMax * upperlimit) * 0.5) / 3.3
required = Abs(required)
'achieved = 0

If TransformerUse > upperlimit Then

For y = 1 To Assign_Profiles.NoEV
    max = 0
    For j = 1 To Assign_Profiles.NoEV
            If max < Charge(j) Then
                If EVFlags(j) = 1 Then
                    max = Charge(j)
                    comp = j
                End If
            End If
    Next
        If achieved < required Then
            If comp > 0 Then
                If EVFlags(comp) = 1 Then
                    EVFlags(comp) = 2
                    achieved = achieved + 1
            
                    Call DisconnectEV(comp)
            
                End If
            End If
    End If
    If achieved = required Then
        Exit For
    End If
Next

ElseIf TransformerUse < lowerlimit Then
For m = 1 To Assign_Profiles.NoEV
    min = 1000
    
    For k = 1 To Assign_Profiles.NoEV
        'If feedercurrents(iter, Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(3, k)) / CheckValues.feedercurrentmax < 0.9 Then
        If MaxLaterals(Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(2, k), Assign_Profiles.EVLocation(3, k)) > LateralsAssigned(Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(2, k), Assign_Profiles.EVLocation(3, k)) Then
            If MaxFeeders(Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(3, k)) > FeedersAssigned(Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(3, k)) Then
                If EVFlags(k) = 2 Then
                    If min > Charge(k) Then
                        min = Charge(k)
                        comp = k
                    End If
                End If
            End If
        End If
    Next

    If achievedcon < required Then
        If comp > 0 Then
            If EVFlags(comp) = 2 Then
                EVFlags(comp) = 1
                achievedcon = achievedcon + 1
                LateralsAssigned(Assign_Profiles.EVLocation(1, comp), Assign_Profiles.EVLocation(2, comp), Assign_Profiles.EVLocation(3, comp)) = LateralsAssigned(Assign_Profiles.EVLocation(1, comp), Assign_Profiles.EVLocation(2, comp), Assign_Profiles.EVLocation(3, comp)) + 1
                FeedersAssigned(Assign_Profiles.EVLocation(1, comp), Assign_Profiles.EVLocation(3, comp)) = FeedersAssigned(Assign_Profiles.EVLocation(1, comp), Assign_Profiles.EVLocation(3, comp)) + 1
                
                Call ConnectEV(comp)
            End If
        End If
    End If

    If achievedcon = required Then
        Exit For
    End If
Next
End If


End Sub

Public Sub CheckEV()

    Dim i As Integer
    Dim Currentss As Variant

    For i = 1 To Assign_Profiles.NoEV
        DSSCircuit.SetActiveElement ("load.EV" & i)
        Currentss = DSSCircuit.ActiveCktElement.Currents
        
        If Currentss(0) <> 0 Then
            EVFlags(i) = 1
        Else
            DSSCircuit.Loads.name = "EV" & i
            If DSSCircuit.Loads.kW = 0 Then
                EVFlags(i) = 2
            Else
                EVFlags(i) = 0
            End If
        End If
        
        If Charge(i) >= MaxCharge(i) Then
            Call DisconnectEV(i)
            EVFlags(i) = 3
        End If
        
        If EVFlags(i) = 1 Then
            Charge(i) = Charge(i) + 1
        End If

    Next

End Sub

Public Sub CalculateDisconnectionsLaterals(ByVal iter As Integer)

    ReDim MaxLaterals(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 3)
    ReDim LateralsAssigned(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 3)
    ReDim MaxFeeders(1 To Assign_Profiles.NoFeeders, 1 To 3)
    ReDim FeedersAssigned(1 To Assign_Profiles.NoFeeders, 1 To 3)
      
    For i = 1 To Assign_Profiles.NoFeeders
        For y = 1 To Assign_Profiles.NoFeeders
            For z = 1 To 3
                
                LateralsAssigned(i, y, z) = 0
                FeedersAssigned(i, z) = 0
                
                If Start.Laterals(iter, i, y, z) < CheckValues.lateralcurrentmax Then
                    
                    MaxLaterals(i, y, z) = (CheckValues.lateralcurrentmax - Start.Laterals(iter, i, y, z)) / (16 * 2)
                Else
                
                    MaxLaterals(i, y, z) = 0
                End If
                
                
                If Start.Feeders(iter, i, z) < (CheckValues.feedercurrentmax) Then
                    
                    MaxFeeders(i, z) = ((CheckValues.feedercurrentmax) - Start.Feeders(iter, i, z)) / (16 * 2.5)
                Else
                
                    MaxFeeders(i, z) = 0
                End If
            
            Next
        Next
    Next

End Sub

Public Sub ConnectEV(ByVal EV)
    
    DSSCircuit.Loads.name = "EV" & EV
    DSSCircuit.Loads.kW = 3.3

End Sub

Public Sub DisconnectEV(ByVal EV)

    DSSCircuit.Loads.name = "EV" & EV
    DSSCircuit.Loads.kW = 0

End Sub
