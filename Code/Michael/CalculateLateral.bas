Attribute VB_Name = "CalculateLateral"
Public Function CustomersPerLateralPerFeeder(noOfCustomers As Long, noOfFeeders As Integer, noOfLaterals As Integer) As Variant
    Dim customersPerFeederPerLateral() As Variant
    Dim customersPerLaterals() As Variant
    Dim noOfCustomersTemp As Integer
    Dim noOfCustomersPerFeeder As Integer
    Dim noOfCustomersPerLateral As Integer
    
    ReDim customersPerFeederPerLateral(1 To noOfFeeders, 1 To noOfLaterals)
    
    noOfCustomersTemp = noOfCustomers
    ' Determine the number of customers per Feeder and Lateral
    noOfCustomersPerFeeder = Int(noOfCustomers / noOfFeeders)
    noOfCustomersPerLateral = Int(noOfCustomersPerFeeder / noOfLaterals)
    
    ' Alocation of customers on each lateral of each feeder
    For i = 1 To noOfFeeders
        'ReDim customersPerLaterals(1 To noOfLaterals)
        For j = 1 To noOfLaterals
            'customersPerLaterals(j) = noOfCustomersPerLateral
            noOfCustomersTemp = noOfCustomersTemp - noOfCustomersPerLateral
            customersPerFeederPerLateral(i, j) = noOfCustomersPerLateral
        Next j
    Next i
    ' If any customers are not allocated, allocate them on the last lateral of each feeder
    Do While noOfCustomersTemp > 0
        i = 1
        Do While noOfCustomersTemp > 0 And i <= noOfFeeders
            customersPerFeederPerLateral(i, noOfLaterals) = customersPerFeederPerLateral(i, noOfLaterals) + 1
            noOfCustomersTemp = noOfCustomersTemp - 1
            i = i + 1
        Loop
    Loop
    
    CustomersPerLateralPerFeeder = customersPerFeederPerLateral
End Function

Public Function PresetLateralSizes() As Variant
    
    Dim LateralSizes(1 To 4, 1 To 4) As Variant
    
    If PresetNetwork.network = "Urban" Then
    
        LateralSizes(1, 1) = 17
        LateralSizes(2, 1) = 17
        LateralSizes(3, 1) = 17
        LateralSizes(4, 1) = 17
        
        LateralSizes(1, 2) = 53
        LateralSizes(2, 2) = 53
        LateralSizes(3, 2) = 53
        LateralSizes(4, 2) = 53
        
        LateralSizes(1, 3) = 44
        LateralSizes(2, 3) = 44
        LateralSizes(3, 3) = 44
        LateralSizes(4, 3) = 44
        
        LateralSizes(1, 4) = 44
        LateralSizes(2, 4) = 44
        LateralSizes(3, 4) = 44
        LateralSizes(4, 4) = 44
        
    
    ElseIf PresetNetwork.network = "SemiUrban" Then
    
        LateralSizes(1, 1) = 12
        LateralSizes(2, 1) = 12
        LateralSizes(3, 1) = 12
        LateralSizes(4, 1) = 12
        
        LateralSizes(1, 2) = 39
        LateralSizes(2, 2) = 39
        LateralSizes(3, 2) = 39
        LateralSizes(4, 2) = 39
        
        LateralSizes(1, 3) = 33
        LateralSizes(2, 3) = 33
        LateralSizes(3, 3) = 33
        LateralSizes(4, 3) = 33
        
        LateralSizes(1, 4) = 33
        LateralSizes(2, 4) = 33
        LateralSizes(3, 4) = 33
        LateralSizes(4, 4) = 33
    
    ElseIf PresetNetwork.network = "Rural" Then
    
        LateralSizes(1, 1) = 4
        LateralSizes(2, 1) = 4
        LateralSizes(3, 1) = 4
        LateralSizes(4, 1) = 4
        
        LateralSizes(1, 2) = 11
        LateralSizes(2, 2) = 11
        LateralSizes(3, 2) = 11
        LateralSizes(4, 2) = 11
        
        LateralSizes(1, 3) = 9
        LateralSizes(2, 3) = 9
        LateralSizes(3, 3) = 9
        LateralSizes(4, 3) = 9
        
        LateralSizes(1, 4) = 9
        LateralSizes(2, 4) = 9
        LateralSizes(3, 4) = 9
        LateralSizes(4, 4) = 9
    End If
        
    PresetLateralSizes = LateralSizes
End Function

Public Function LateralNo(ByRef customer As Integer) As Integer

    If PresetNetwork.network = "Urban" Then
    
        If customer <= 17 Then
            LateralNo = 1
        ElseIf customer <= 53 + 17 Then
            LateralNo = 2
        ElseIf customer <= 53 + 17 + 44 Then
            LateralNo = 3
        Else
            LateralNo = 4
        End If
    
    ElseIf PresetNetwork.network = "SemiUrban" Then
        
        If customer <= 12 Then
            LateralNo = 1
        ElseIf customer <= 39 + 12 Then
            LateralNo = 2
        ElseIf customer <= 39 + 33 + 12 Then
            LateralNo = 3
        Else
            LateralNo = 4
        End If
    
    ElseIf PresetNetwork.network = "Rural" Then
        
        If customer <= 4 Then
            LateralNo = 1
        ElseIf customer <= 4 + 11 Then
            LateralNo = 2
        ElseIf customer <= 4 + 11 + 9 Then
            LateralNo = 3
        Else
            LateralNo = 4
        End If
    End If

End Function

Public Function feederLength(ByVal LateralNo As Integer) As Integer

    If PresetNetwork.network = "Urban" Then
    
        If LateralNo = 1 Then feederLength = 35
        If LateralNo = 2 Then feederLength = 35 + 69
        If LateralNo = 3 Or LateralNo = 4 Then feederLength = 35 + 69 + 70
    
    ElseIf PresetNetwork.network = "SemiUrban" Then
        
        If LateralNo = 1 Then feederLength = 47
        If LateralNo = 2 Then feederLength = 47 + 94
        If LateralNo = 3 Or LateralNo = 4 Then feederLength = 47 + 94 + 94
    
    ElseIf PresetNetwork.network = "Rural" Then
        
        If LateralNo = 1 Then feederLength = 50
        If LateralNo = 2 Then feederLength = 150
        If LateralNo = 3 Or LateralNo = 4 Then feederLength = 250

    End If
    

End Function

Public Function lateralLength(ByVal LateralNo As Integer, ByVal CustomerNo As Integer) As Integer
    
    If PresetNetwork.network = "Urban" Then
        
        If LateralNo = 1 Then
            lateralLength = Int(136 * CustomerNo / 17)
        ElseIf LateralNo = 2 Then
            lateralLength = Int(136 * (CustomerNo - 17) / 53)
        ElseIf LateralNo = 3 Or LateralNo = 4 Then
            lateralLength = Int(136 * (CustomerNo - 17 - 53) / 44)
        ElseIf LateralNo = 4 Then
            lateralLength = Int(136 * (CustomerNo - 17 - 53 - 44) / 44)
        End If
        
    ElseIf PresetNetwork.network = "SemiUrban" Then
        
        If LateralNo = 1 Then
            lateralLength = Int(185 * CustomerNo / 12)
        ElseIf LateralNo = 2 Then
            lateralLength = Int(185 * (CustomerNo - 12) / 39)
        ElseIf LateralNo = 3 Or LateralNo = 4 Then
            lateralLength = Int(185 * (CustomerNo - 12 - 39) / 33)
        ElseIf LateralNo = 4 Then
            lateralLength = Int(185 * (CustomerNo - 12 - 39 - 33) / 44)
        End If
    
    ElseIf PresetNetwork.network = "Rural" Then
        
        If LateralNo = 1 Then
            lateralLength = Int(196 * CustomerNo / 4)
        ElseIf LateralNo = 2 Then
            lateralLength = Int(196 * (CustomerNo - 4) / 11)
        ElseIf LateralNo = 3 Or LateralNo = 4 Then
            lateralLength = Int(196 * (CustomerNo - 4 - 11) / 9)
        ElseIf LateralNo = 4 Then
            lateralLength = Int(196 * (CustomerNo - 4 - 11 - 9) / 9)
        End If

    End If

End Function

Function LateralLocation(ByVal LateralNo As Integer, ByVal CustomerNo As Integer) As Integer

    If PresetNetwork.network = "Urban" Then
        
        If LateralNo = 1 Then
            If (CustomerNo / 17) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        
        ElseIf LateralNo = 2 Then

            If ((CustomerNo - 17) / 53) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        ElseIf LateralNo = 3 Or LateralNo = 4 Then

            If ((CustomerNo - 17 - 53) / 44) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        ElseIf LateralNo = 4 Then

            If ((CustomerNo - 17 - 53 - 44) / 44) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        End If
        
    ElseIf PresetNetwork.network = "SemiUrban" Then
        
        If LateralNo = 1 Then

            If (CustomerNo / 12) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        ElseIf LateralNo = 2 Then

            If ((CustomerNo - 12) / 39) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        ElseIf LateralNo = 3 Or LateralNo = 4 Then

            If ((CustomerNo - 12 - 39) / 33) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        ElseIf LateralNo = 4 Then

            If ((CustomerNo - 12 - 39 - 33) / 44) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        End If
    
    ElseIf PresetNetwork.network = "Rural" Then
        
        If LateralNo = 1 Then

            If (CustomerNo / 4) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        ElseIf LateralNo = 2 Then

            If ((CustomerNo - 4) / 11) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        ElseIf LateralNo = 3 Or LateralNo = 4 Then

            If ((CustomerNo - 4 - 11) / 9) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        ElseIf LateralNo = 4 Then

            If ((CustomerNo - 4 - 11 - 9) / 9) > 0.5 Then
                LateralLocation = 2
            Else
                LateralLocation = 1
            End If
        End If

    End If

End Function
