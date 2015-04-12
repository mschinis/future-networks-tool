Attribute VB_Name = "ANMhp"
Public HPReduction() As Double
Public TransformerFlagHP As Boolean
Public FeederFlagHP As Boolean
Public LateralFlagHP As Boolean
Public HPReductionArray() As Double
Public HPIncreaseArray() As Double
Public HPFlags() As Integer

Public AchievedFeederHP() As Double
Public AchievedHP As Double

Public MaxLateralsHP() As Integer
Public LateralsAssignedHP() As Integer
Public MaxFeedersHP() As Integer
Public FeedersAssignedHP() As Integer


Public Sub HPManagement(ByVal i As Integer)


    ReDim HPReductionArray(1 To Assign_Profiles.NoHP)
    ReDim HPIncreaseArray(1 To Assign_Profiles.NoHP)
    HPReduction = GetHeatPumps17
    Call CheckHP(i)
    Call ManageDisconnectionsHP(i)
    Call CalculateDisconnectionsLateralsHP(i)

'    Call LateralManagementHP(i, Start.Laterals)
'    Call FeederManagementHP(i, Start.Feeders)
    Call TransformerManagementHP(i, Start.TransformerArray(i, 1) / CheckValues.TransformerMax)

End Sub

Public Sub LateralManagementHP(ByVal iter As Integer, ByRef CurrentUse() As Double)

'Dim i, y, z, h As Integer
'Dim lateralrequired As Integer
'
''ReDim achievedlaterals(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 3)
'ReDim AchievedFeedersHP(1 To Assign_Profiles.NoLaterals, 1 To 3)
'AchievedHP = 0
'
'For i = 1 To Assign_Profiles.NoFeeders
'    For y = 1 To Assign_Profiles.NoLaterals
'        For z = 1 To 3
'            If CurrentUse(iter, i, y, z) / CheckValues.lateralcurrentmax > 1 Then
'                For h = 1 To Assign_Profiles.NoEV
'                    'max = 0
'                    lateralrequired = (CurrentUse(iter, i, y, z) - CheckValues.lateralcurrentmax) * 0.5 / 16
''                    For A = 1 To Assign_Profiles.NoEV
''                        If Assign_Profiles.EVLocation(1, A) = i And Assign_Profiles.EVLocation(2, A) = y And Assign_Profiles.EVLocation(3, A) = z Then
''                            If EVFlags(A) = 1 Then
''                                If max < Charge(A) Then
''                                    max = Charge(A)
''                                    comp = A
''                                End If
''                            End If
''                        End If
''                    Next
'
'                    If achievedlaterals(i, y, z) < lateralrequired Then
'                        If comp > 0 Then
'                            If EVFlags(comp) = 1 Then
'                                EVFlags(comp) = 4
'                                achieved = achieved + 1
'                                achievedlaterals(i, y, z) = achievedlaterals(i, y, z) + 1
'                                achievedfeeders(i, z) = achievedfeeders(i, z) + 1
'                                Call DisconnectEV(comp)
'                            End If
'                        End If
'                    End If
'                    If achievedlaterals(i, y, z) = lateralrequired Then
'                        Exit For
'                    End If
'                Next
'            End If
'        Next
'    Next
'Next
'
'
'

End Sub

Public Sub FeederManagementHP(ByVal iter As Integer, ByRef CurrentUse() As Double)

Dim i, y, z, h As Integer
Dim feederrequired As Double

'ReDim achievedlaterals(1 To 4, 1 To 4, 1 To 3)
ReDim AchievedFeederHP(1 To 4, 1 To 3)
AchievedHP = 0

For i = 1 To Assign_Profiles.NoFeeders
        For z = 1 To 3
        feederrequired = (CurrentUse(iter, i, z) - CheckValues.feedercurrentmax * 0.95)
        feederrequired = (feederrequired * 250) / 2000
            
            If CurrentUse(iter, i, z) / CheckValues.feedercurrentmax > 0.95 Then
                For h = 1 To Assign_Profiles.NoEV
'                    max = 0

'                    For A = 1 To Assign_Profiles.NoEV
'                        If Assign_Profiles.EVLocation(1, A) = i And Assign_Profiles.EVLocation(3, A) = z Then
'                            If EVFlags(A) = 1 Then
'                                If max < Charge(A) Then
'                                    max = Charge(A)
'                                    comp = A
'                                End If
'                            End If
'                        End If
'                    Next
                
                    If AchievedFeederHP(i, z) < feederrequired Then
                        If Assign_Profiles.HPLocation(1, h) = i And Assign_Profiles.HPLocation(3, h) = z Then
                            If HPFlags(h) = 1 Then
                                EVFlags(h) = 2
                                AchievedHP = AchievedHP + HPReductionArray(h)
                                achievedfeeders(i, z) = achievedfeeders(i, z) + HPReductionArray(h)
                                Call DisconnectHP(h, iter)
                            End If
                        End If
                    End If
                    If achievedfeeders(i, z) > feederrequired Then
                        Exit For
                    End If
                Next
            End If
        Next
Next


End Sub

Public Sub TransformerManagementHP(ByVal iter As Integer, ByVal TransformerUse As Double)

Dim RequiredDisc, RequiredCon As Double
Dim max, min, comp, j, y, i, m, k As Integer
Dim upperlimit, lowerlimit As Double

Dim AchievedHPCon As Double

AchievedHPCon = 0

upperlimit = 0.96
lowerlimit = 0.93

RequiredDisc = Abs(((TransformerUse * CheckValues.TransformerMax) - CheckValues.TransformerMax * upperlimit)) / 2
RequiredCon = Abs(((TransformerUse * CheckValues.TransformerMax) - CheckValues.TransformerMax * lowerlimit)) / 2

AchievedHP = 0

If TransformerUse > upperlimit And HPReduction(iter) < 1 Then

For y = 1 To Assign_Profiles.NoHP
'    max = 0
'    For j = 1 To Assign_Profiles.NoEV
'            If max < Charge(j) Then
'                If EVFlags(j) = 1 Then
'                    max = Charge(j)
'                    comp = j
'                End If
'            End If
'    Next



    If AchievedHP < RequiredDisc Then
        If HPFlags(y) = 1 Then
            HPFlags(y) = 2
            AchievedHP = AchievedHP + HPReductionArray(y)
            
            Call DisconnectHP(y, iter)
        End If

    End If
    If AchievedHP > RequiredDisc Then
        Exit For
    End If
Next

ElseIf TransformerUse < lowerlimit Then
For k = 1 To Assign_Profiles.NoHP
'    min = 1000
    
'    For k = 1 To Assign_Profiles.NoEV
'        'If feedercurrents(iter, Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(3, k)) / CheckValues.feedercurrentmax < 0.9 Then
'        If MaxLaterals(Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(2, k), Assign_Profiles.EVLocation(3, k)) > LateralsAssigned(Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(2, k), Assign_Profiles.EVLocation(3, k)) Then
'            If MaxFeeders(Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(3, k)) > FeedersAssigned(Assign_Profiles.EVLocation(1, k), Assign_Profiles.EVLocation(3, k)) Then
'                If EVFlags(k) = 2 Then
'                    If min > Charge(k) Then
'                        min = Charge(k)
'                        comp = k
'                    End If
'                End If
'            End If
'        End If
'    Next

    If AchievedHPCon < RequiredCon Then

            If HPFlags(k) = 2 Then
                If MaxFeedersHP(Assign_Profiles.HPLocation(1, k), Assign_Profiles.HPLocation(3, k)) > FeedersAssignedHP(Assign_Profiles.HPLocation(1, k), Assign_Profiles.HPLocation(3, k)) Then
                    HPFlags(k) = 1
                    AchievedHPCon = AchievedHPCon + HPIncreaseArray(k)
                    LateralsAssignedHP(Assign_Profiles.HPLocation(1, k), Assign_Profiles.HPLocation(2, k), Assign_Profiles.HPLocation(3, k)) = LateralsAssignedHP(Assign_Profiles.HPLocation(1, k), Assign_Profiles.HPLocation(2, k), Assign_Profiles.HPLocation(3, k)) + 1
                    FeedersAssignedHP(Assign_Profiles.HPLocation(1, k), Assign_Profiles.HPLocation(3, k)) = FeedersAssignedHP(Assign_Profiles.HPLocation(1, k), Assign_Profiles.HPLocation(3, k)) + 1
                
                    Call ConnectHP(k)
                End If
            End If

    End If

    If AchievedHPCon > required Then
        Exit For
    End If
Next
End If

End Sub

Function CheckHP(ByVal iter As Integer)

    Dim i As Integer
    Dim Powerss As Variant
    Dim TempPower As Double
    Dim TempReduction As Double
    Dim TempIncrease As Double
    If iter > 600 Then
        upperlimit = upperlimit
    End If
    For i = 1 To Assign_Profiles.NoHP

        DSSCircuit.SetActiveElement ("load.HP" & i)
        Powerss = DSSCircuit.ActiveCktElement.Powers
        TempPower = Sqr((Powerss(0) ^ 2) + (Powerss(1) ^ 2))

        If HPFlags(i) = 1 Then
            If HPReduction(iter) < 1 Then
                TempReduction = TempPower * HPReduction(iter)
                TempReduction = TempPower - TempReduction
            Else
                TempReduction = 0
            End If

        End If

        HPReductionArray(i) = TempReduction

    Next
    
'If iter > 700 Then
'    TempReduction = TempReduction
'End If

End Function

Function ManageDisconnectionsHP(ByVal iter As Integer)
    
    Dim Powerss As Variant
    Dim TempPower As Double
    Dim TempIncrease As Double
    
    For i = 1 To Assign_Profiles.NoHP
        If HPFlags(i) = 2 Then
            
            DSSCircuit.SetActiveElement ("load.HP" & i)
            Powerss = DSSCircuit.ActiveCktElement.Powers
            TempPower = Sqr((Powerss(0) ^ 2) + (Powerss(1) ^ 2))
            
            If HPReduction(iter) < 1 Then
                DSSCircuit.Loads.name = "HP" & i
                DSSCircuit.Loads.kW = Round(1 * HPReduction(iter), 2)
                TempIncrease = TempPower / HPReduction(iter)
                TempIncrease = TempPower - TempIncrease
            Else
                
                DSSCircuit.Loads.name = "HP" & i
                DSSCircuit.Loads.kW = 1
                HPFlags(i) = 1
                TempIncrease = 0
            End If
        
            HPIncreaseArray(i) = TempIncrease
        
        End If
    Next

End Function

Function CalculateDisconnectionsLateralsHP(ByVal iter As Integer)

    ReDim MaxLateralsHP(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 3)
    ReDim LateralsAssignedHP(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 3)
    ReDim MaxFeedersHP(1 To Assign_Profiles.NoFeeders, 1 To 3)
    ReDim FeedersAssignedHP(1 To Assign_Profiles.NoFeeders, 1 To 3)
      
    For i = 1 To Assign_Profiles.NoFeeders
        For y = 1 To Assign_Profiles.NoFeeders
            For z = 1 To 3
                
                LateralsAssignedHP(i, y, z) = 0
                FeedersAssignedHP(i, z) = 0
                
                If Start.Laterals(iter, i, y, z) < CheckValues.lateralcurrentmax Then
                    
                    MaxLateralsHP(i, y, z) = (CheckValues.lateralcurrentmax - Start.Laterals(iter, i, y, z)) * 240 / 2000
                Else
                
                    MaxLateralsHP(i, y, z) = 0
                End If
                
                
                If Start.Feeders(iter, i, z) < (CheckValues.feedercurrentmax) Then
                    
                    MaxFeedersHP(i, z) = ((CheckValues.feedercurrentmax) - Start.Feeders(iter, i, z)) * 240 / 2000
                Else
                
                    MaxFeedersHP(i, z) = 0
                End If
            
            Next
        Next
    Next

End Function

Function DisconnectHP(ByVal HP As Integer, ByVal iter As Integer)

    DSSCircuit.Loads.name = "HP" & HP
   DSSCircuit.Loads.kW = Round(1 * HPReduction(iter), 2)

    
End Function

Function ConnectHP(ByVal HP As Integer)
    
    DSSCircuit.Loads.name = "HP" & HP
    DSSCircuit.Loads.kW = 1

End Function

Function GetHeatPumps17() As Variant


    Dim dataArray() As String
    Dim DoubleArray() As Double
    Dim endloop As Integer
    Dim i As Integer

    ReDim DoubleArray(1 To 1440)

    Dim strFileName As String
    strFileName = ThisWorkbook.Path & "/Loadshapes/HP/HeatPumps17.txt"
    Open strFileName For Input As #1

     ' -------- read from txt file to dataArrayay -------- '

     i = 0
     Do Until EOF(1)
        ReDim Preserve dataArray(i)
        Line Input #1, dataArray(i)
        i = i + 1
     Loop
     Close #1

    endloop = i

    For i = 1 To endloop
        DoubleArray(i) = CDbl(dataArray(i - 1))
    Next
    GetHeatPumps17 = DoubleArray
End Function
