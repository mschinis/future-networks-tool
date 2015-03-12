Attribute VB_Name = "CalculateLateral"
Public Function PresetLateralSizes() As Variant
    
    Dim lateralsizes(1 To 4, 1 To 4) As Integer
    
    If PresetNetwork.Network = "Urban" Then
    
        lateralsizes(1, 1) = 17
        lateralsizes(2, 1) = 17
        lateralsizes(3, 1) = 17
        lateralsizes(4, 1) = 17
        
        lateralsizes(1, 2) = 53
        lateralsizes(2, 2) = 53
        lateralsizes(3, 2) = 53
        lateralsizes(4, 2) = 53
        
        lateralsizes(1, 3) = 44
        lateralsizes(2, 3) = 44
        lateralsizes(3, 3) = 44
        lateralsizes(4, 3) = 44
        
        lateralsizes(1, 4) = 44
        lateralsizes(2, 4) = 44
        lateralsizes(3, 4) = 44
        lateralsizes(4, 4) = 44
        
    
    ElseIf PresetNetwork.Network = "SemiUrban" Then
    
        lateralsizes(1, 1) = 12
        lateralsizes(2, 1) = 12
        lateralsizes(3, 1) = 12
        lateralsizes(4, 1) = 12
        
        lateralsizes(1, 2) = 39
        lateralsizes(2, 2) = 39
        lateralsizes(3, 2) = 39
        lateralsizes(4, 2) = 39
        
        lateralsizes(1, 3) = 33
        lateralsizes(2, 3) = 33
        lateralsizes(3, 3) = 33
        lateralsizes(4, 3) = 33
        
        lateralsizes(1, 4) = 33
        lateralsizes(2, 4) = 33
        lateralsizes(3, 4) = 33
        lateralsizes(4, 4) = 33
        
    
    ElseIf PresetNetwork.Network = "Rural" Then
        
        lateralsizes(1, 1) = 4
        lateralsizes(2, 1) = 4
        lateralsizes(3, 1) = 4
        lateralsizes(4, 1) = 4
        
        lateralsizes(1, 2) = 11
        lateralsizes(2, 2) = 11
        lateralsizes(3, 2) = 11
        lateralsizes(4, 2) = 11
        
        lateralsizes(1, 3) = 9
        lateralsizes(2, 3) = 9
        lateralsizes(3, 3) = 9
        lateralsizes(4, 3) = 9
        
        lateralsizes(1, 4) = 9
        lateralsizes(2, 4) = 9
        lateralsizes(3, 4) = 9
        lateralsizes(4, 4) = 9
        
    End If
    
        PresetLateralSizes = lateralsizes
    
    
End Function

Public Function LateralNo(ByRef customer As Integer) As Integer

    If PresetNetwork.Network = "Urban" Then
    
        If customer <= 17 Then
            LateralNo = 1
        ElseIf customer <= 53 + 17 Then
            LateralNo = 2
        ElseIf customer <= 53 + 17 + 44 Then
            LateralNo = 3
        Else
            LateralNo = 4
        End If
    
    ElseIf PresetNetwork.Network = "SemiUrban" Then
        
        If customer <= 12 Then
            LateralNo = 1
        ElseIf customer <= 39 + 12 Then
            LateralNo = 2
        ElseIf customer <= 39 + 33 + 12 Then
            LateralNo = 3
        Else
            LateralNo = 4
        End If
    
    ElseIf PresetNetwork.Network = "Rural" Then
        
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

    If PresetNetwork.Network = "Urban" Then
    
        If LateralNo = 1 Then feederLength = 35
        If LateralNo = 2 Then feederLength = 35 + 69
        If LateralNo = 3 Or LateralNo = 4 Then feederLength = 35 + 69 + 70
    
    ElseIf PresetNetwork.Network = "SemiUrban" Then
        
        If LateralNo = 1 Then feederLength = 47
        If LateralNo = 2 Then feederLength = 47 + 94
        If LateralNo = 3 Or LateralNo = 4 Then feederLength = 47 + 94 + 94
    
    ElseIf PresetNetwork.Network = "Rural" Then
        
        If LateralNo = 1 Then feederLength = 50
        If LateralNo = 2 Then feederLength = 150
        If LateralNo = 3 Or LateralNo = 4 Then feederLength = 250

    End If
    

End Function

Public Function lateralLength(ByVal LateralNo As Integer, ByVal CustomerNo As Integer) As Integer
    
    If PresetNetwork.Network = "Urban" Then
        
        If LateralNo = 1 Then
            lateralLength = Int(136 * CustomerNo / 17)
        ElseIf LateralNo = 2 Then
            lateralLength = Int(136 * (CustomerNo - 17) / 53)
        ElseIf LateralNo = 3 Or LateralNo = 4 Then
            lateralLength = Int(136 * (CustomerNo - 17 - 53) / 44)
        ElseIf LateralNo = 4 Then
            lateralLength = Int(136 * (CustomerNo - 17 - 53 - 44) / 44)
        End If
        
    ElseIf PresetNetwork.Network = "SemiUrban" Then
        
        If LateralNo = 1 Then
            lateralLength = Int(185 * CustomerNo / 12)
        ElseIf LateralNo = 2 Then
            lateralLength = Int(185 * (CustomerNo - 12) / 39)
        ElseIf LateralNo = 3 Or LateralNo = 4 Then
            lateralLength = Int(185 * (CustomerNo - 12 - 39) / 33)
        ElseIf LateralNo = 4 Then
            lateralLength = Int(185 * (CustomerNo - 12 - 39 - 33) / 44)
        End If
    
    ElseIf PresetNetwork.Network = "Rural" Then
        
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

    If PresetNetwork.Network = "Urban" Then
        
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
        
    ElseIf PresetNetwork.Network = "SemiUrban" Then
        
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
    
    ElseIf PresetNetwork.Network = "Rural" Then
        
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
