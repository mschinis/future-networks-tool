Attribute VB_Name = "ANMpv"
Public PVFlags() As Integer
Public EnergyReceived() As Integer

Public achievedPV As Integer
Public achievedlateralsPV() As Integer
Public achievedfeedersPV() As Integer

Public reachievedPV As Integer
Public reachievedlateralsPV() As Integer
Public reachievedfeedersPV() As Integer

Public spointPV As Integer
Public previousdisc As Integer
Public requiredsaved() As Integer




Public Sub PVManagement(ByVal i As Integer)

        Call Reconnect
        Call LateralManagementPV(i, Start.Laterals)
        
End Sub

Public Sub LateralManagementPV(ByVal iter As Integer, ByRef CurrentUse() As Double)

Dim i, y, z, h As Integer
Dim lateralrequired As Integer
Dim limit As Double
Dim gain As Integer

'If ChooseNetwork.TransformerTap = 0 Then limit = 1.085
'If ChooseNetwork.TransformerTap = -2.5 Then limit = 1.075
'If ChooseNetwork.TransformerTap = -5 Then limit = 1.065

gain = 3000

If ((Start.TransformerArray(iter, 2) + Start.TransformerArray(iter, 3) + Start.TransformerArray(iter, 4)) / 3) > 1.07 Then
    
    limit = 1.085

    If (Assign_Profiles.NoPV / PresetNetwork.customers) < 0.5 Then limit = 1.09
    
ElseIf ((Start.TransformerArray(iter, 2) + Start.TransformerArray(iter, 3) + Start.TransformerArray(iter, 4)) / 3) > 1.05 Then
     
     limit = 1.078

     If (Assign_Profiles.NoPV / PresetNetwork.customers) < 0.5 Then limit = 1.08
     
ElseIf ((Start.TransformerArray(iter, 2) + Start.TransformerArray(iter, 3) + Start.TransformerArray(iter, 4)) / 3) > 1.03 Then
     
     limit = 1.075

     If (Assign_Profiles.NoPV / PresetNetwork.customers) < 0.5 Then limit = 1.075
Else
    limit = 1.075
End If




ReDim achievedlateralsPV(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 3)
ReDim achievedfeedersPV(1 To Assign_Profiles.NoLaterals, 1 To 3)
achievedPV = 0

ReDim reachievedlateralsPV(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 3)
ReDim reachievedfeedersPV(1 To Assign_Profiles.NoLaterals, 1 To 3)
reachievedPV = 0

For i = 1 To Assign_Profiles.NoFeeders
    For y = 1 To Assign_Profiles.NoLaterals
        For z = 1 To 3

            lateralrequired = ((CurrentUse(iter, i, y, z + 6) - limit) * gain)
            If lateralrequired < 0 Then lateralrequired = Int(lateralrequired / 3)
            lateralrequired = lateralrequired + requiredsaved(i, z)

            If lateralrequired > 0 Then
                InternalIter = 0
                Do While achievedfeedersPV(i, z) < lateralrequired

                    If PVFlags(spointPV) = 1 And Assign_Profiles.PVLocation(1, spointPV) = i And Assign_Profiles.PVLocation(3, spointPV) = z Then
                        Call DisconnectPV(spointPV)

                        distancevariable = Assign_Profiles.PVLocation(2, spointPV)
                        If Assign_Profiles.PVLocation(2, spointPV) = 4 Then distancevariable = 3
                        distancevariable = distancevariable + Assign_Profiles.PVLocation(6, spointPV)
                        achievedfeedersPV(i, z) = achievedfeedersPV(i, z) + Assign_Profiles.PVLocation(4, spointPV) + distancevariable
                        achievedPV = achievedPV + Assign_Profiles.PVLocation(4, spointPV) + distancevariable
                        PVFlags(spointPV) = 2
                    End If

                    spointPV = 1 + spointPV
                    If spointPV Mod Assign_Profiles.NoPV = 1 Then spointPV = 1

                    InternalIter = InternalIter + 1
                    If InternalIter = Assign_Profiles.NoPV Then
                        Exit Do
                    End If

                Loop
            requiredsaved(i, z) = achievedfeedersPV(i, z)
            End If

        Next
    Next
Next



End Sub

Public Sub Reconnect()

    For i = 1 To Assign_Profiles.NoPV
        
        Call ConnectPV(i)
        PVFlags(i) = 1
    Next


End Sub

Public Sub DisconnectPV(ByVal PV As Integer)

    DSSCircuit.Generators.name = "PV" & PV
    DSSCircuit.Generators.kW = 0


End Sub

Public Sub ConnectPV(ByVal PV As Integer)
    
    DSSCircuit.Generators.name = "PV" & PV
    DSSCircuit.Generators.kW = 10

End Sub
