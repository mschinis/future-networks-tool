Attribute VB_Name = "Assign_Profiles"
Public CustomersArrayHP() As Variant
Public HPStopPoint() As Integer
Public CHPStopPoint() As Integer
Public HouseStopPoint() As Integer
Public house As Integer

Public HPEnabled As Boolean
Public CHPEnabled As Boolean

Public PVLocation() As Integer
Public EVLocation() As Integer
Public HPLocation() As Integer
Public NoPV As Integer
Public NoEV As Integer
Public NoHP As Integer

Public NoFeeders As Integer
Public NoLaterals As Integer

Public LateralSizes() As Integer

Public PVPenetrationArray() As Double
Public EVPenetrationArray() As Double
Public HPPenetrationArray() As Double
Public CHPPenetrationArray() As Double



Function Assign_PV_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double, ByVal location As Integer, ByVal Tmonth As Integer, ByVal clearness As Integer)

Dim LoadshapeNumber As Integer
Dim PVsize As Integer
Dim CustomersArray() As Variant


ReDim ANMpv.PVFlags(1 To NoCustomers)
ReDim ANMpv.requiredsaved(1 To NoFeeders, 1 To 3)
'Dim LateralSizes() As Integer
Dim TempArray() As Variant
Dim PenetrationNumberDouble As Double
Dim PenetrationNumberInteger As Integer
Dim PenetrationNumber As Integer
Dim PenetrationPercentage As Integer
Dim PenetrationMatrix(1 To 100) As Integer
Dim DevicesNumber As Integer


'If PresetNetwork.Network = "Urban" Or PresetNetwork.Network = "SemiUrban" Or PresetNetwork.Network = "Rural" Then
'    LateralSizes = PresetLateralSizes
'End If
DevicesNumber = 0
max = 0
For i = 1 To NoFeeders
    For y = 1 To NoLaterals
        
        If PVPenetrationArray(i, y) = Empty Then PVPenetrationArray(i, y) = penetration ' PUT THIS INTO AN IF STATEMENT LATER
        If max < LateralSizes(i, y) Then
            max = LateralSizes(i, y)
        End If
        DevicesNumber = DevicesNumber + (LateralSizes(i, y) * PVPenetrationArray(i, y) + 1)
    Next
Next

ReDim CustomersArray(1 To NoFeeders, 1 To NoLaterals, 1 To max)
ReDim PVLocation(1 To 6, 1 To DevicesNumber)

For i = 1 To NoFeeders
    h = 0
    For y = 1 To NoLaterals
        ReDim TempArray(1 To LateralSizes(i, y))
        For z = 1 To LateralSizes(i, y)
            h = h + 1
            If LateralSizes(i, y) <> 0 Then
                TempArray(z) = i & "_" & h
            End If
        Next
        
        TempArray = ShuffleArray(TempArray)
        For z = 1 To LateralSizes(i, y)
            CustomersArray(i, y, z) = TempArray(z)
        Next
    Next
Next

DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\PV"

m = 0
For i = 1 To NoFeeders
    For y = 1 To NoLaterals
        
        PenetrationNumberDouble = LateralSizes(i, y) * PVPenetrationArray(i, y)
        PenetrationNumberInteger = LateralSizes(i, y) * PVPenetrationArray(i, y)
        PenetrationPercentage = (PenetrationNumberDouble - PenetrationNumberInteger) * 100
        
        For u = 1 To 100
            If PenetrationPercentage < 1 Then
                If u <= Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger - 1
                ElseIf u > Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger
                End If
            End If
            If PenetrationPercentage > 1 Then
                If u <= Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger + 1
                ElseIf u > Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger
                End If
            End If
            If PenetrationPercentage = 0 Then PenetrationMatrix(u) = PenetrationNumberInteger
            
        Next
        
        For z = 1 To PenetrationMatrix(Int((100 - 1 + 1) * Rnd + 1))
            PVsize = Int((4 - 1 + 1) * Rnd + 1)
                m = m + 1
                DSSText.Command = "new loadshape.PVload" & m & " npts=1440 minterval=1.0 csvfile=PV" & location & "_" & Tmonth & "_" & clearness & "_" & PVsize & ".txt"
                DSSText.Command = "new generator.PV" & m & " bus1=Consumer" & CustomersArray(i, y, z) & ".1 Phases=1 kV=0.23 kW=10 PF=1 Daily=PVload" & m
                
                PVLocation(1, m) = i 'Store the feeder of the device
                PVLocation(3, m) = Int(Mid(CustomersArray(i, y, z), 3)) Mod 3 'Store the phase of each device
                If PVLocation(3, m) = 0 Then PVLocation(3, m) = 3
                PVLocation(2, m) = y 'Store the lateral of each device

                PVLocation(4, m) = PVsize

                PVLocation(5, m) = feederLength(PVLocation(2, m))
                PVLocation(6, m) = LateralLocation(PVLocation(2, m), Int(Mid(CustomersArray(i, y, z), 3)))

                ANMpv.PVFlags(m) = 1
        Next
    Next
Next


ANMpv.spointPV = 1
ANMpv.previousdisc = 0
NoPV = m
End Function

Function Assign_House_Profiles(ByVal NoCustomers As Integer, ByVal Tmonth As Integer, ByVal Tday As Integer)

Dim LoadshapeNumber As Integer
Dim OccupantsArray() As Integer
ReDim OccupantsArray(1 To NoCustomers)
Dim occupants As Integer
'Dim LateralSizes() As Integer

'If PresetNetwork.Network = "Urban" Or PresetNetwork.Network = "SemiUrban" Or PresetNetwork.Network = "Rural" Then
'    LateralSizes = PresetLateralSizes
'End If

If HPEnabled = False And CHPEnabled = False Then
    max = 0
    For i = 1 To NoFeeders
        For y = 1 To NoLaterals
            If max < LateralSizes(i, y) Then
                max = LateralSizes(i, y)
            End If
        Next
    Next
    
    house = 0
    ReDim HouseStopPoint(1 To NoFeeders, 1 To NoLaterals)
    ReDim CustomersArrayHP(1 To NoFeeders, 1 To NoLaterals, 1 To max)
    For i = 1 To NoFeeders
        h = 0
        For y = 1 To NoLaterals
            ReDim TempArray(1 To LateralSizes(i, y))
            For z = 1 To LateralSizes(i, y)
                h = h + 1
                If LateralSizes(i, y) <> 0 Then
                    TempArray(z) = i & "_" & h
                End If
            Next
            
            TempArray = ShuffleArray(TempArray)
            For z = 1 To LateralSizes(i, y)
                CustomersArrayHP(i, y, z) = TempArray(z)
            Next
        Next
    Next
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

For i = 1 To NoFeeders
    For y = 1 To NoLaterals
        For z = HouseStopPoint(i, y) + 1 To LateralSizes(i, y)
            
            house = house + 1
            LoadshapeNumber = Int((200 - 1 + 1) * Rnd + 1)
            occupants = OccupantsArray(Int((100 - 1 + 1) * Rnd + 1))
            
            DSSText.Command = "new loadshape.Houseload" & house & " npts=1440 minterval=1.0 csvfile=House" & Tmonth & "_" & Tday & "_" & occupants & "_" & LoadshapeNumber & "_1.txt"
            DSSText.Command = "new load.House" & house & " bus1=Consumer" & CustomersArrayHP(i, y, z) & ".1 Phases=1 kV=0.23 kW=10 PF=0.97 Daily=Houseload" & house
            
        Next
    Next
Next

End Function

Function Assign_HP_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double, ByVal Tmonth As Integer, ByVal Tday As Integer, ByVal location As Integer)

HPEnabled = True

Dim HouseTypeArray(1 To 100) As Integer
Dim InsulationTypeArray(1 To 100) As Integer
Dim occupants As Integer
Dim OccupantsArray() As Integer
ReDim OccupantsArray(1 To NoCustomers)
'Dim LateralSizes() As Integer
Dim TempArray() As Variant
Dim PenetrationNumberDouble As Double
Dim PenetrationNumberInteger As Integer
Dim PenetrationNumber As Integer
Dim PenetrationPercentage As Integer
Dim PenetrationMatrix(1 To 100) As Integer
Dim DevicesNumber As Integer


'---------------------------------- Create probability arrays ----------------------------
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

'--------------------------- Location and month mapping -------------------------------------------

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
'------------------------------------------------------------------------------------




'If PresetNetwork.Network = "Urban" Or PresetNetwork.Network = "SemiUrban" Or PresetNetwork.Network = "Rural" Then
'    LateralSizes = PresetLateralSizes
'End If

max = 0
For i = 1 To NoFeeders
    For y = 1 To NoLaterals
       If HPPenetrationArray(i, y) = Empty Then HPPenetrationArray(i, y) = penetration ' PUT THIS INTO AN IF STATEMENT LATER
        If max < LateralSizes(i, y) Then
            max = LateralSizes(i, y)
        End If
        DevicesNumber = DevicesNumber + (LateralSizes(i, y) * HPPenetrationArray(i, y) + 1)
    Next
Next

ReDim HPLocation(1 To 3, 1 To DevicesNumber)
ReDim CustomersArrayHP(1 To NoFeeders, 1 To NoLaterals, 1 To max)

For i = 1 To NoFeeders
    h = 0
    For y = 1 To NoLaterals
        ReDim TempArray(1 To LateralSizes(i, y))
        For z = 1 To LateralSizes(i, y)
            h = h + 1
            If LateralSizes(i, y) <> 0 Then
                TempArray(z) = i & "_" & h
            End If
        Next
        
        TempArray = ShuffleArray(TempArray)
        For z = 1 To LateralSizes(i, y)
            CustomersArrayHP(i, y, z) = TempArray(z)
        Next
    Next
Next

house = 0
ReDim HouseStopPoint(1 To NoFeeders, 1 To NoLaterals)
m = 0

For i = 1 To NoFeeders
    For y = 1 To NoLaterals
        
        PenetrationNumberDouble = LateralSizes(i, y) * HPPenetrationArray(i, y)
        PenetrationNumberInteger = LateralSizes(i, y) * HPPenetrationArray(i, y)
        PenetrationPercentage = (PenetrationNumberDouble - PenetrationNumberInteger) * 100
        
        For u = 1 To 100
            If PenetrationPercentage < 0 Then
                If u <= Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger - 1
                ElseIf u > Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger
                End If
            End If
            If PenetrationPercentage > 0 Then
                If u <= Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger + 1
                ElseIf u > Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger
                End If
            End If
            If PenetrationPercentage = 0 Then PenetrationMatrix(u) = PenetrationNumberInteger
            
        Next
        
        For z = 1 To PenetrationMatrix(Int((100 - 1 + 1) * Rnd + 1))
            
                m = m + 1
                house = house + 1
                HouseStopPoint(i, y) = HouseStopPoint(i, y) + 1
                
                repetition = Int((20 - 1 + 1) * Rnd + 1)
                Thouse = HouseTypeArray(Int((100 - 1 + 1) * Rnd + 1))
                Tinsulation = InsulationTypeArray(Int((100 - 1 + 1) * Rnd + 1))
                occupants = OccupantsArray(Int((100 - 1 + 1) * Rnd + 1))

                DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\HP"
                DSSText.Command = "new loadshape.HPload" & m & " npts=1440 minterval=1.0 csvfile=HP" & TmonthAdj & "_" & Tday & "_" & location & "_" & Thouse & "_" & Tinsulation & "_" & occupants & "_" & repetition & ".txt"
                DSSText.Command = "new load.HP" & m & " bus1=Consumer" & CustomersArrayHP(i, y, z) & ".1 Phases=1 kV=0.23 kW=1 PF=0.9 Daily=HPload" & m
            
                DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\House"
                LoadshapeNumber = Int((500 - 1 + 1) * Rnd + 1)
                DSSText.Command = "new loadshape.Houseload" & m & " npts=1440 minterval=1.0 csvfile=House" & Tmonth & "_" & Tday & "_" & occupants & "_" & LoadshapeNumber & ".txt"
                DSSText.Command = "new load.House" & m & " bus1=Consumer" & CustomersArrayHP(i, y, z) & ".1 Phases=1 kV=0.23 kW=10 PF=0.97 Daily=Houseload" & m
            
                HPLocation(1, m) = i 'Store the feeder of the device
                HPLocation(3, m) = Int(Mid(CustomersArrayHP(i, y, z), 3)) Mod 3 'Store the phase of each device
                If HPLocation(3, m) = 0 Then HPLocation(3, m) = 3
                HPLocation(2, m) = y 'Store the lateral of each device

        Next
        HPStopPoint(i, y) = z
    Next
Next

NoHP = m


'For i = 1 To 4
'    For y = 1 To NoCustomers / 4
'        z = 1 + z
'        CustomersArray(z) = i & "_" & y
'    Next
'Next
'
'
'CustomersArrayShuffledHP = ShuffleArray(CustomersArray)
'
'
'
'For i = 1 To (NoCustomers) * penetration
'
'
'    repetition = Int((20 - 1 + 1) * Rnd + 1)
'    Thouse = HouseTypeArray(Int((100 - 1 + 1) * Rnd + 1))
'    Tinsulation = InsulationTypeArray(Int((100 - 1 + 1) * Rnd + 1))
'    occupants = OccupantsArray(Int((100 - 1 + 1) * Rnd + 1))
'
'    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\HP"
'    DSSText.Command = "new loadshape.HPload" & i & " npts=1440 minterval=1.0 csvfile=HP" & TmonthAdj & "_" & Tday & "_" & location & "_" & Thouse & "_" & Tinsulation & "_" & occupants & "_" & repetition & ".txt"
'    DSSText.Command = "new load.HP" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=1 PF=0.9 Daily=HPload" & i
'
'    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\House"
'    LoadshapeNumber = Int((500 - 1 + 1) * Rnd + 1)
'    DSSText.Command = "new loadshape.Houseload" & i & " npts=1440 minterval=1.0 csvfile=House" & Tmonth & "_" & Tday & "_" & occupants & "_" & LoadshapeNumber & ".txt"
'    DSSText.Command = "new load.House" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=10 PF=0.97 Daily=Houseload" & i
'
'    HPLocation(1, i) = Int(Left(CustomersArrayShuffledHP(i), 1)) 'Store the feeder of the device
'    HPLocation(3, i) = Int(Mid(CustomersArrayShuffledHP(i), 3)) Mod 3 'Store the phase of each device
'    If HPLocation(3, i) = 0 Then HPLocation(3, i) = 3
'    HPLocation(2, i) = LateralNo(Int(Mid(CustomersArrayShuffledHP(i), 3))) 'Store the lateral of each device
'
'Next
'
'HPStopPoint = (NoCustomers * penetration)


End Function

Function Assign_CHP_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double, ByVal Tmonth As Integer, ByVal Tday As Integer, ByVal location As Integer)

CHPEnabled = True
Dim CustomersArray() As Variant
Dim HouseTypeArray(1 To 100) As Integer
Dim InsulationTypeArray(1 To 100) As Integer
Dim occupants As Integer
Dim OccupantsArray() As Integer
ReDim OccupantsArray(1 To NoCustomers)
'Dim LateralSizes() As Integer
Dim TempArray() As Variant
Dim PenetrationNumberDouble As Double
Dim PenetrationNumberInteger As Integer
Dim PenetrationNumber As Integer
Dim PenetrationPercentage As Integer
Dim PenetrationMatrix(1 To 100) As Integer
Dim DevicesNumber As Integer


'---------------------------------- Create probability arrays ----------------------------
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

'--------------------------- Location and month mapping -------------------------------------------

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
'------------------------------------------------------------------------------------




'If PresetNetwork.Network = "Urban" Or PresetNetwork.Network = "SemiUrban" Or PresetNetwork.Network = "Rural" Then
'    LateralSizes = PresetLateralSizes
'End If

max = 0
For i = 1 To NoFeeders
    For y = 1 To NoLaterals
       If CHPPenetrationArray(i, y) = Empty Then CHPPenetrationArray(i, y) = penetration ' PUT THIS INTO AN IF STATEMENT LATER
        If max < LateralSizes(i, y) Then
            max = LateralSizes(i, y)
        End If
        DevicesNumber = DevicesNumber + (LateralSizes(i, y) * CHPPenetrationArray(i, y) + 1)
    Next
Next

ReDim CHPLocation(1 To 3, 1 To DevicesNumber)

If HPEnabled = False Then
    ReDim CustomersArrayHP(1 To NoFeeders, 1 To NoLaterals, 1 To max)
    
    For i = 1 To NoFeeders
        h = 0
        For y = 1 To NoLaterals
            ReDim TempArray(1 To LateralSizes(i, y))
            For z = 1 To LateralSizes(i, y)
                h = h + 1
                If LateralSizes(i, y) <> 0 Then
                    TempArray(z) = i & "_" & h
                End If
            Next
            
            TempArray = ShuffleArray(TempArray)
            For z = 1 To LateralSizes(i, y)
                CustomersArrayHP(i, y, z) = TempArray(z)
            Next
        Next
    Next
    house = 0
    ReDim HouseStopPoint(1 To NoFeeders, 1 To NoLaterals)
End If
m = 0


For i = 1 To NoFeeders
    For y = 1 To NoLaterals
        
        PenetrationNumberDouble = LateralSizes(i, y) * CHPPenetrationArray(i, y)
        PenetrationNumberInteger = LateralSizes(i, y) * CHPPenetrationArray(i, y)
        PenetrationPercentage = (PenetrationNumberDouble - PenetrationNumberInteger) * 100
        
        For u = 1 To 100
            If PenetrationPercentage < 1 Then
                If u <= Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger - 1
                ElseIf u > Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger
                End If
            End If
            If PenetrationPercentage > 1 Then
                If u <= Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger + 1
                ElseIf u > Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger
                End If
            End If
            If PenetrationPercentage = 0 Then PenetrationMatrix(u) = PenetrationNumberInteger
            
        Next
        
        For z = HPStopPoint(i, y) + 1 To PenetrationMatrix(Int((100 - 1 + 1) * Rnd + 1)) + HPStopPoint(i, y)
            
                If z <= LateralSizes(i, y) Then
                    m = m + 1
                    house = house + 1
                    HouseStopPoint(i, y) = HouseStopPoint(i, y) + 1
                    
                    repetition = Int((20 - 1 + 1) * Rnd + 1)
                    Thouse = HouseTypeArray(Int((100 - 1 + 1) * Rnd + 1))
                    Tinsulation = InsulationTypeArray(Int((100 - 1 + 1) * Rnd + 1))
                    occupants = OccupantsArray(Int((100 - 1 + 1) * Rnd + 1))
            
                    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\CHP"
                    DSSText.Command = "new loadshape.CHPload" & m & " npts=1440 minterval=1.0 csvfile=CHP" & TmonthAdj & "_" & Tday & "_" & location & "_" & Thouse & "_" & Tinsulation & "_" & occupants & "_" & repetition & ".txt"
                    DSSText.Command = "new generator.CHP" & m & " bus1=Consumer" & CustomersArrayHP(i, y, z) & ".1 Phases=1 kV=0.23 kW=1 PF=1 Daily=CHPload" & m
            
                    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\House"
                    LoadshapeNumber = Int((500 - 1 + 1) * Rnd + 1)
                    DSSText.Command = "new loadshape.Houseload" & house & " npts=1440 minterval=1.0 csvfile=House" & Tmonth & "_" & Tday & "_" & occupants & "_" & LoadshapeNumber & ".txt"
                    DSSText.Command = "new load.House" & house & " bus1=Consumer" & CustomersArrayHP(i, y, z) & ".1 Phases=1 kV=0.23 kW=10 PF=0.97 Daily=Houseload" & house
                End If
        Next
        CHPStopPoint(i, y) = z
    Next
Next

NoCHP = m




'Dim CustomersArray() As Variant
'ReDim CustomersArray(1 To NoCustomers)
'Dim InsulationTypeArray(1 To 100) As Integer
'Dim HouseTypeArray(1 To 100) As Integer
'Dim occupants As Integer
'Dim OccupantsArray() As Integer
'ReDim OccupantsArray(1 To NoCustomers)
'
'
'For i = 1 To 100
'    If i <= 19 Then
'        InsulationTypeArray(i) = 1
'    ElseIf i <= 19 + 44 Then
'        InsulationTypeArray(i) = 2
'    Else
'        InsulationTypeArray(i) = 3
'    End If
'Next
'
'If Tmonth >= 1 And Tmonth <= 2 Then TmonthAdj = 1
'If Tmonth = 12 Then TmonthAdj = 1
'If Tmonth >= 3 And Tmonth <= 5 Then TmonthAdj = 2
'If Tmonth >= 9 And Tmonth <= 11 Then TmonthAdj = 2
'If Tmonth >= 6 And Tmonth <= 8 Then TmonthAdj = 3
'
'For i = 1 To 100
'
'    If i <= 30 Then
'        OccupantsArray(i) = 1
'    ElseIf i <= 65 Then
'        OccupantsArray(i) = 2
'    ElseIf i <= 80 Then
'        OccupantsArray(i) = 3
'    ElseIf i <= 93 Then
'        OccupantsArray(i) = 4
'    Else
'        OccupantsArray(i) = 5
'    End If
'
'Next
'
'For i = 1 To 100
'    If i <= 25 Then
'        HouseTypeArray(i) = 1
'    ElseIf i <= 25 + 27 Then
'        HouseTypeArray(i) = 2
'    ElseIf i <= 25 + 27 + 30 Then
'        HouseTypeArray(i) = 3
'    Else
'        HouseTypeArray(i) = 4
'    End If
'Next
'
'
'If location = 2 Or location = 3 Then
'    location = 2
'ElseIf location = 4 Or location = 5 Or location = 6 Or location = 7 Or location = 8 Then
'    location = 4
'ElseIf location = 9 Or location = 10 Or location = 11 Then
'    location = 4
'End If
'
'
'If HPStopPoint = 0 Then
'
'    For i = 1 To 4
'        For y = 1 To NoCustomers / 4
'            z = 1 + z
'            CustomersArray(z) = i & "_" & y
'        Next
'    Next
'
'
'CustomersArrayShuffledHP = ShuffleArray(CustomersArray)
'
'End If
'
'
'
'For i = (HPStopPoint + 1) To ((NoCustomers) * penetration) + HPStopPoint
'
'
'    repetition = Int((20 - 1 + 1) * Rnd + 1)
'    Thouse = HouseTypeArray(Int((100 - 1 + 1) * Rnd + 1))
'    Tinsulation = InsulationTypeArray(Int((100 - 1 + 1) * Rnd + 1))
'    occupants = OccupantsArray(Int((100 - 1 + 1) * Rnd + 1))
'
'    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\CHP"
'    DSSText.Command = "new loadshape.CHPload" & i & " npts=1440 minterval=1.0 csvfile=CHP" & TmonthAdj & "_" & Tday & "_" & location & "_" & Thouse & "_" & Tinsulation & "_" & occupants & "_" & repetition & ".txt"
'    DSSText.Command = "new generator.CHP" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=1 PF=1 Daily=CHPload" & i
'
'    DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\House"
'    LoadshapeNumber = Int((500 - 1 + 1) * Rnd + 1)
'    DSSText.Command = "new loadshape.Houseload" & i & " npts=1440 minterval=1.0 csvfile=House" & Tmonth & "_" & Tday & "_" & occupants & "_" & LoadshapeNumber & ".txt"
'    DSSText.Command = "new load.House" & i & " bus1=Consumer" & CustomersArrayShuffledHP(i) & ".1 Phases=1 kV=0.23 kW=10 PF=0.97 Daily=Houseload" & i
'
'Next
'
'CHPStopPoint = NoCustomers * penetration

End Function

Function Assign_EV_Profiles(ByVal NoCustomers As Integer, ByVal penetration As Double)

Dim LoadshapeNumber, dis As Integer
Dim CustomersArray() As Variant

'Dim LateralSizes() As Integer
Dim TempArray() As Variant
Dim PenetrationNumberDouble As Double
Dim PenetrationNumberInteger As Integer
Dim PenetrationNumber As Integer
Dim PenetrationPercentage As Integer
Dim PenetrationMatrix(1 To 100) As Integer


'If PresetNetwork.Network = "Urban" Or PresetNetwork.Network = "SemiUrban" Or PresetNetwork.Network = "Rural" Then
'    LateralSizes = PresetLateralSizes
'End If

max = 0
For i = 1 To NoFeeders
    For y = 1 To NoLaterals
       If EVPenetrationArray(i, y) = Empty Then EVPenetrationArray(i, y) = penetration ' PUT THIS INTO AN IF STATEMENT LATER
        If max < LateralSizes(i, y) Then
            max = LateralSizes(i, y)
        End If
        DevicesNumber = DevicesNumber + (LateralSizes(i, y) * EVPenetrationArray(i, y) + 1)
    Next
Next

ReDim CustomersArray(1 To NoFeeders, 1 To NoLaterals, 1 To max)

For i = 1 To NoFeeders
    h = 0
    For y = 1 To NoLaterals
        ReDim TempArray(1 To LateralSizes(i, y))
        For z = 1 To LateralSizes(i, y)
            h = h + 1
            If LateralSizes(i, y) <> 0 Then
                TempArray(z) = i & "_" & h
            End If
        Next
        
        TempArray = ShuffleArray(TempArray)
        For z = 1 To LateralSizes(i, y)
            CustomersArray(i, y, z) = TempArray(z)
        Next
    Next
Next

DSSText.Command = "set Datapath=" & ActiveWorkbook.Path & "\Loadshapes\EV"
ReDim EVLocation(1 To 3, 1 To DevicesNumber)

m = 0
For i = 1 To NoFeeders
    For y = 1 To NoLaterals
        
        PenetrationNumberDouble = LateralSizes(i, y) * EVPenetrationArray(i, y)
        PenetrationNumberInteger = LateralSizes(i, y) * EVPenetrationArray(i, y)
        PenetrationPercentage = (PenetrationNumberDouble - PenetrationNumberInteger) * 100
        
        For u = 1 To 100
            If PenetrationPercentage < 1 Then
                If u <= Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger - 1
                ElseIf u > Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger
                End If
            End If
            If PenetrationPercentage > 1 Then
                If u <= Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger + 1
                ElseIf u > Abs(PenetrationPercentage) Then
                    PenetrationMatrix(u) = PenetrationNumberInteger
                End If
            End If
            If PenetrationPercentage = 0 Then PenetrationMatrix(u) = PenetrationNumberInteger
            
        Next
        
        For z = 1 To PenetrationMatrix(Int((100 - 1 + 1) * Rnd + 1))

            m = m + 1
            LoadshapeNumber = Int((1000 - 1 + 1) * Rnd + 1)

            DSSText.Command = "new loadshape.EVload" & m & " npts=1440 minterval=1.0 csvfile=EV" & LoadshapeNumber & ".txt"
            DSSText.Command = "new load.EV" & m & " bus1=Consumer" & CustomersArray(i, y, z) & ".1 Phases=1 kV=0.23 kW=3.3 PF=1 Daily=EVload" & m

            EVLocation(1, m) = i 'Store the feeder of the device
            EVLocation(3, m) = Int(Mid(CustomersArray(i, y, z), 3)) Mod 3 'Store the phase of each device
            If EVLocation(3, m) = 0 Then EVLocation(3, m) = 3
            EVLocation(2, m) = y 'Store the lateral of each device
        Next
    Next
Next

NoEV = m

ReDim ANMev.Charge(1 To NoEV)
ReDim ANMev.MaxCharge(1 To NoEV)
ReDim ANMev.EVFlags(1 To NoEV)


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


End Function

Function ShuffleArray(InArray() As Variant) As Variant()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShuffleArray
' This function returns the values of InArray in random order. The original
' InArray is not modified.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim l As Long
    Dim temp As Variant
    Dim j As Long
    Dim arr() As Variant
    
    
    Randomize
    l = UBound(InArray) - LBound(InArray) + 1
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
