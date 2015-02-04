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
