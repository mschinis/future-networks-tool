Attribute VB_Name = "Assign_Profiles"
Public CustomersArrayShuffledHP() As Variant
Public HPStopPoint As Integer
Public CHPStopPoint As Integer


Function Assign_PV_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double, ByVal location As Integer, ByVal Tmonth As Integer, ByVal clearness As Integer)

Dim LoadshapeNumber As Integer
Dim PVsize As Integer
Dim CustomersArray() As Variant
ReDim CustomersArray(1 To NoCustomers)


For i = 1 To 4
    For y = 1 To NoCustomers / 4
        Z = 1 + Z
        CustomersArray(Z) = i & "_" & y
    Next
Next


CustomersArrayShuffled = ShuffleArray(CustomersArray)

DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\PV"

For i = 1 To (penetration * NoCustomers)

    PVsize = Int((4 - 1 + 1) * Rnd + 1)

    
    DSSText.Command = "new loadshape.PVload" & i & " npts=1440 minterval=1.0 csvfile=PV" & location & "_" & Tmonth & "_" & clearness & "_" & PVsize & ".txt"
    DSSText.Command = "new generator.PV" & i & " bus1=Consumer" & CustomersArrayShuffled(i) & ".1 Phases=1 kV=0.23 kW=10 PF=1 Daily=PVload" & i


Next

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
            Z = 1 + Z
            CustomersArray(Z) = i & "_" & y
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
        Z = 1 + Z
        CustomersArray(Z) = i & "_" & y
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
            Z = 1 + Z
            CustomersArray(Z) = i & "_" & y
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

Dim LoadshapeNumber As Integer
Dim CustomersArray() As Variant
ReDim CustomersArray(1 To NoCustomers)


For i = 1 To 4
    For y = 1 To NoCustomers / 4
        Z = 1 + Z
        CustomersArray(Z) = i & "_" & y
    Next
Next


CustomersArrayShuffled = ShuffleArray(CustomersArray)

DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\EV"

For i = 1 To (penetration * NoCustomers)


    LoadshapeNumber = Int((1000 - 1 + 1) * Rnd + 1)
    
    DSSText.Command = "new loadshape.EVload" & i & " npts=1440 minterval=1.0 csvfile=EV" & LoadshapeNumber & ".txt"
    DSSText.Command = "new load.EV" & i & " bus1=Consumer" & CustomersArrayShuffled(i) & ".1 Phases=1 kV=0.23 kW=3.3 PF=1 Daily=EVload" & i


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






