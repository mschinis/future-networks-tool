Attribute VB_Name = "Assign_Profiles"
Public CustomersArrayShuffledHP() As Variant
Public HPStopPoint As Integer
Public CHPStopPoint As Integer

Public PVLocation() As Integer
Public EVLocation() As Integer
Public HPLocation() As Integer
Public NoPV As Integer
Public NoEV As Integer
Public NoHP As Integer




Function Assign_PV_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double, ByVal location As Integer, ByVal Tmonth As Integer, ByVal clearness As Integer)

Dim LoadshapeNumber As Integer
Dim PVsize As Integer
Dim CustomersArray() As Variant
ReDim CustomersArray(1 To NoCustomers)
ReDim PVLocation(1 To 6, 1 To NoCustomers * penetration)
NoPV = NoCustomers * penetration
ReDim ANMpv.PVFlags(1 To NoCustomers)
ReDim ANMpv.requiredsaved(1 To 4, 1 To 3)

For i = 1 To 4
    For y = 1 To NoCustomers / 4
        z = 1 + z
        CustomersArray(z) = i & "_" & y
    Next
Next


customersarrayshuffled = ShuffleArray(CustomersArray)

DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\PV"

For i = 1 To (penetration * NoCustomers)

    PVsize = Int((4 - 1 + 1) * Rnd + 1)

    
    DSSText.Command = "new loadshape.PVload" & i & " npts=1440 minterval=1.0 csvfile=PV" & location & "_" & Tmonth & "_" & clearness & "_" & PVsize & ".txt"
    DSSText.Command = "new generator.PV" & i & " bus1=Consumer" & customersarrayshuffled(i) & ".1 Phases=1 kV=0.23 kW=10 PF=1 Daily=PVload" & i

    
    PVLocation(1, i) = Int(Left(customersarrayshuffled(i), 1)) 'Store the feeder of the device
    PVLocation(3, i) = Int(Mid(customersarrayshuffled(i), 3)) Mod 3 'Store the phase of each device
    If PVLocation(3, i) = 0 Then PVLocation(3, i) = 3
    PVLocation(2, i) = LateralNo(Int(Mid(customersarrayshuffled(i), 3))) 'Store the lateral of each device
    
    PVLocation(4, i) = PVsize
    
    PVLocation(5, i) = FeederLength(PVLocation(2, i))
    PVLocation(6, i) = LateralLocation(PVLocation(2, i), Int(Mid(customersarrayshuffled(i), 3)))
    
    ANMpv.PVFlags(i) = 1
    
Next

ANMpv.spointPV = 1
ANMpv.previousdisc = 0
End Function

Function Assign_House_Profiles(ByVal NoCustomers As Integer, ByVal Tmonth As Integer, ByVal Tday As Integer)

Dim LoadshapeNumber As Integer
Dim CustomersArray() As Variant
ReDim CustomersArray(1 To NoCustomers)
Dim OccupantsArray() As Integer
ReDim OccupantsArray(1 To NoCustomers)
Dim occupants As Integer

If HPStopPoint = 0 And CHPStopPoint = 0 Then
    For i = 1 To 4
        For y = 1 To NoCustomers / 4
            z = 1 + z
            CustomersArray(z) = i & "_" & y
        Next
    Next
    CustomersArrayShuffledHP = ShuffleArray(CustomersArray)
End If

For i = 1 To 100

    If i <= 30 Then
        OccupantsArray(i) = 1
    ElseIf i <= 65 Then
        OccupantsArray(i) = 2
    ElseIf i <= 80 Then
        OccupantsArray(i) = 3
    ElseIf i <= 93 Then
        OccupantsArray(i) = 4
    Else
        OccupantsArray(i) = 5
    End If

Next


DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\House"

For i = HPStopPoint + CHPStopPoint + 1 To (NoCustomers)


    LoadshapeNumber = Int((200 - 1 + 1) * Rnd + 1)
    occupants = OccupantsArray(Int((100 - 1 + 1) * Rnd + 1))
    

    DSSText.Command = "new loadshape.Houseload" & i & " npts=1440 minterval=1.0 csvfile=House" & Tmonth & "_" & Tday & "_" & occupants & "_" & LoadshapeNumber & "_1.txt"
    DSSText.Command = "new load.House" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=10 PF=0.97 Daily=Houseload" & i


Next

End Function

Function Assign_HP_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double, ByVal Tmonth As Integer, ByVal Tday As Integer, ByVal location As Integer)

Dim CustomersArray() As Variant
ReDim CustomersArray(1 To NoCustomers)
Dim HouseTypeArray(1 To 100) As Integer
Dim InsulationTypeArray(1 To 100) As Integer
Dim occupants As Integer
Dim OccupantsArray() As Integer
ReDim OccupantsArray(1 To NoCustomers)
ReDim HPLocation(1 To 3, 1 To NoCustomers * penetration)
NoHP = NoCustomers * penetration


For i = 1 To 100
    If i <= 19 Then
        InsulationTypeArray(i) = 1
    ElseIf i <= 19 + 44 Then
        InsulationTypeArray(i) = 2
    Else
        InsulationTypeArray(i) = 3
    End If
Next

For i = 1 To 100
    If i <= 25 Then
        HouseTypeArray(i) = 1
    ElseIf i <= 25 + 27 Then
        HouseTypeArray(i) = 2
    ElseIf i <= 25 + 27 + 30 Then
        HouseTypeArray(i) = 3
    Else
        HouseTypeArray(i) = 4
    End If
Next
For i = 1 To 100

    If i <= 30 Then
        OccupantsArray(i) = 1
    ElseIf i <= 65 Then
        OccupantsArray(i) = 2
    ElseIf i <= 80 Then
        OccupantsArray(i) = 3
    ElseIf i <= 93 Then
        OccupantsArray(i) = 4
    Else
        OccupantsArray(i) = 5
    End If

Next



If Tmonth >= 1 And Tmonth <= 2 Then TmonthAdj = 1
If Tmonth = 12 Then TmonthAdj = 1
If Tmonth >= 3 And Tmonth <= 5 Then TmonthAdj = 2
If Tmonth >= 9 And Tmonth <= 11 Then TmonthAdj = 2
If Tmonth >= 6 And Tmonth <= 8 Then TmonthAdj = 3

If location = 2 Or location = 3 Then
    location = 2
ElseIf location = 4 Or location = 5 Or location = 6 Or location = 7 Or location = 8 Then
    location = 4
ElseIf location = 9 Or location = 10 Or location = 11 Then
    location = 4
End If


For i = 1 To 4
    For y = 1 To NoCustomers / 4
        z = 1 + z
        CustomersArray(z) = i & "_" & y
    Next
Next


CustomersArrayShuffledHP = ShuffleArray(CustomersArray)



For i = 1 To (NoCustomers) * penetration


    repetition = Int((20 - 1 + 1) * Rnd + 1)
    Thouse = HouseTypeArray(Int((100 - 1 + 1) * Rnd + 1))
    Tinsulation = InsulationTypeArray(Int((100 - 1 + 1) * Rnd + 1))
    occupants = OccupantsArray(Int((100 - 1 + 1) * Rnd + 1))
    
    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\HP"
    DSSText.Command = "new loadshape.HPload" & i & " npts=1440 minterval=1.0 csvfile=HP" & TmonthAdj & "_" & Tday & "_" & location & "_" & Thouse & "_" & Tinsulation & "_" & occupants & "_" & repetition & ".txt"
    DSSText.Command = "new load.HP" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=1 PF=0.9 Daily=HPload" & i
    
    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\House"
    LoadshapeNumber = Int((500 - 1 + 1) * Rnd + 1)
    DSSText.Command = "new loadshape.Houseload" & i & " npts=1440 minterval=1.0 csvfile=House" & Tmonth & "_" & Tday & "_" & occupants & "_" & LoadshapeNumber & ".txt"
    DSSText.Command = "new load.House" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=10 PF=0.97 Daily=Houseload" & i
    
    HPLocation(1, i) = Int(Left(CustomersArrayShuffledHP(i), 1)) 'Store the feeder of the device
    HPLocation(3, i) = Int(Mid(CustomersArrayShuffledHP(i), 3)) Mod 3 'Store the phase of each device
    If HPLocation(3, i) = 0 Then HPLocation(3, i) = 3
    HPLocation(2, i) = LateralNo(Int(Mid(CustomersArrayShuffledHP(i), 3))) 'Store the lateral of each device
    
Next

HPStopPoint = (NoCustomers * penetration)


End Function

Function Assign_CHP_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double, ByVal Tmonth As Integer, ByVal Tday As Integer, ByVal location As Integer)

Dim CustomersArray() As Variant
ReDim CustomersArray(1 To NoCustomers)
Dim InsulationTypeArray(1 To 100) As Integer
Dim HouseTypeArray(1 To 100) As Integer
Dim occupants As Integer
Dim OccupantsArray() As Integer
ReDim OccupantsArray(1 To NoCustomers)


For i = 1 To 100
    If i <= 19 Then
        InsulationTypeArray(i) = 1
    ElseIf i <= 19 + 44 Then
        InsulationTypeArray(i) = 2
    Else
        InsulationTypeArray(i) = 3
    End If
Next

If Tmonth >= 1 And Tmonth <= 2 Then TmonthAdj = 1
If Tmonth = 12 Then TmonthAdj = 1
If Tmonth >= 3 And Tmonth <= 5 Then TmonthAdj = 2
If Tmonth >= 9 And Tmonth <= 11 Then TmonthAdj = 2
If Tmonth >= 6 And Tmonth <= 8 Then TmonthAdj = 3

For i = 1 To 100

    If i <= 30 Then
        OccupantsArray(i) = 1
    ElseIf i <= 65 Then
        OccupantsArray(i) = 2
    ElseIf i <= 80 Then
        OccupantsArray(i) = 3
    ElseIf i <= 93 Then
        OccupantsArray(i) = 4
    Else
        OccupantsArray(i) = 5
    End If

Next

For i = 1 To 100
    If i <= 25 Then
        HouseTypeArray(i) = 1
    ElseIf i <= 25 + 27 Then
        HouseTypeArray(i) = 2
    ElseIf i <= 25 + 27 + 30 Then
        HouseTypeArray(i) = 3
    Else
        HouseTypeArray(i) = 4
    End If
Next


If location = 2 Or location = 3 Then
    location = 2
ElseIf location = 4 Or location = 5 Or location = 6 Or location = 7 Or location = 8 Then
    location = 4
ElseIf location = 9 Or location = 10 Or location = 11 Then
    location = 4
End If


If HPStopPoint = 0 Then

    For i = 1 To 4
        For y = 1 To NoCustomers / 4
            z = 1 + z
            CustomersArray(z) = i & "_" & y
        Next
    Next


CustomersArrayShuffledHP = ShuffleArray(CustomersArray)

End If



For i = (HPStopPoint + 1) To ((NoCustomers) * penetration) + HPStopPoint


    repetition = Int((20 - 1 + 1) * Rnd + 1)
    Thouse = HouseTypeArray(Int((100 - 1 + 1) * Rnd + 1))
    Tinsulation = InsulationTypeArray(Int((100 - 1 + 1) * Rnd + 1))
    occupants = OccupantsArray(Int((100 - 1 + 1) * Rnd + 1))

    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\CHP"
    DSSText.Command = "new loadshape.CHPload" & i & " npts=1440 minterval=1.0 csvfile=CHP" & TmonthAdj & "_" & Tday & "_" & location & "_" & Thouse & "_" & Tinsulation & "_" & occupants & "_" & repetition & ".txt"
    DSSText.Command = "new generator.CHP" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=1 PF=1 Daily=CHPload" & i
    
    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\House"
    LoadshapeNumber = Int((500 - 1 + 1) * Rnd + 1)
    DSSText.Command = "new loadshape.Houseload" & i & " npts=1440 minterval=1.0 csvfile=House" & Tmonth & "_" & Tday & "_" & occupants & "_" & LoadshapeNumber & ".txt"
    DSSText.Command = "new load.House" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=10 PF=0.97 Daily=Houseload" & i
    
Next

CHPStopPoint = NoCustomers * penetration

End Function

Function Assign_EV_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double)

NoEV = NoCustomers * penetration

Dim LoadshapeNumber, dis As Integer
Dim CustomersArray() As Variant
ReDim CustomersArray(1 To NoCustomers)
ReDim EVLocation(1 To 3, 1 To NoEV)
ReDim ANMev.Charge(1 To NoEV)
ReDim ANMev.MaxCharge(1 To NoEV)
ReDim ANMev.EVFlags(1 To NoEV)

For i = 1 To 4
    For y = 1 To NoCustomers / 4
        z = 1 + z
        CustomersArray(z) = i & "_" & y
    Next
Next

For y = 1 To NoEV
    
    ANMev.EVFlags(y) = 5
    ANMev.Charge(y) = 0
    
'''''''''''''''''''' Create the charge required for each vehicle, based on standard deviation '''''''''''''''''''''''''''''''''''''

    dis = Application.WorksheetFunction.NormInv(Rnd(), 180, 70)
    
    If dis < 1 Or dis > 480 Then
        Do While dis < 1 Or dis > 480
            dis = Application.WorksheetFunction.NormInv(Rnd(), 180, 70)
        Loop
    End If
        
    ANMev.MaxCharge(y) = dis

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next

customersarrayshuffled = ShuffleArray(CustomersArray)

DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\EV"

For i = 1 To NoEV


    LoadshapeNumber = Int((1000 - 1 + 1) * Rnd + 1)
    
    DSSText.Command = "new loadshape.EVload" & i & " npts=1440 minterval=1.0 csvfile=EV" & LoadshapeNumber & ".txt"
    DSSText.Command = "new load.EV" & i & " bus1=Consumer" & customersarrayshuffled(i) & ".1 Phases=1 kV=0.23 kW=3.3 PF=1 Daily=EVload" & i
    
    EVLocation(1, i) = Int(Left(customersarrayshuffled(i), 1)) 'Store the feeder of the device
    EVLocation(3, i) = Int(Mid(customersarrayshuffled(i), 3)) Mod 3 'Store the phase of each device
    If EVLocation(3, i) = 0 Then EVLocation(3, i) = 3
    EVLocation(2, i) = LateralNo(Int(Mid(customersarrayshuffled(i), 3))) 'Store the lateral of each device

Next

End Function

Function ShuffleArray(InArray() As Variant) As Variant()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShuffleArray
' This function returns the values of InArray in random order. The original
' InArray is not modified.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim L As Long
    Dim temp As Variant
    Dim j As Long
    Dim arr() As Variant
    
    
    Randomize
    L = UBound(InArray) - LBound(InArray) + 1
    ReDim arr(LBound(InArray) To UBound(InArray))
    For N = LBound(InArray) To UBound(InArray)
        arr(N) = InArray(N)
    Next N
    For N = LBound(InArray) To UBound(InArray)
        j = Int((UBound(InArray) - LBound(InArray) + 1) * Rnd + LBound(InArray))
        If N <> j Then
            temp = arr(N)
            arr(N) = arr(j)
            arr(j) = temp
        End If
    Next N
    ShuffleArray = arr
End Function






