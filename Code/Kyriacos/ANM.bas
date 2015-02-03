Attribute VB_Name = "ANM"
Public MaxCharge() As Integer
Public Charge() As Integer
Public EVFlags() As Integer
Public achieved As Integer
Public required As Integer


Public Sub EVManagement(ByVal i As Integer)

    Call CheckEV
    
    'if ANMEV = true then    'if ANM for EVs is enabled
    Call TransformerManagementEV(Start.TransformerArray(i, 1) / CheckValues.TransformerMax)
    
    'end if
    
End Sub

Public Sub TransformerManagementEV(ByVal TransformerUse As Double)

Dim max, min, comp, j, y, i, m, k As Integer
Dim bandgap As Double

bandgap = 0.97

required = (((TransformerUse * CheckValues.TransformerMax) - CheckValues.TransformerMax) * 0.5) / 3.3
required = Abs(required)
achieved = 0

If TransformerUse > 1 Then

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

ElseIf TransformerUse < bandgap Then
For m = 1 To Assign_Profiles.NoEV
    min = 1000
    
    For k = 1 To Assign_Profiles.NoEV

            If EVFlags(k) = 2 Then
                If min > Charge(k) Then
                    min = Charge(k)
                    comp = k
                End If
            End If
    Next

    If achieved < required Then
        If comp > 0 Then
            If EVFlags(comp) = 2 Then
                EVFlags(comp) = 1
                achieved = achieved + 1
            
                Call ConnectEV(comp)
            End If
        End If
    End If

    If achieved = required Then
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

Public Sub ConnectEV(ByVal EV)
    
    DSSCircuit.Loads.name = "EV" & EV
    DSSCircuit.Loads.kW = 3.3

End Sub

Public Sub DisconnectEV(ByVal EV)

    DSSCircuit.Loads.name = "EV" & EV
    DSSCircuit.Loads.kW = 0

End Sub
