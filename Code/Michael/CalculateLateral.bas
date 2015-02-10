Attribute VB_Name = "CalculateLateral"
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
