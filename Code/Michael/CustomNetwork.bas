Attribute VB_Name = "CustomNetwork"
Public Sub Custom_Network()





End Sub

Function Create_Transformer(ByVal capacity As Integer, ByVal impedance As Double)

    DSSText.Command = "New transformer.LV_Transformer Buses=(Sourcebus, Main_Busbar)  Conns=(Delta, Wye) kvs=(11, 0.4) kvas=(" & capacity & "," & capacity & ") xhl=" & impedance
    
End Function

Function Create_Linecodes(ByVal R0 As Double, ByVal R1 As Double, ByVal X0 As Double, ByVal X1 As Double, ByVal C0 As Double, ByVal C1 As Double, ByVal phases As Integer, ByVal name As String)

    DSSText.Command = "New Linecode." & name & " R1=" & R1 & " X1=" & X1 & " R0=" & R0 & " X0=" & X0 & " C0=" & C0 & " C1=" & C1 & " units=km nphases=" & phases

End Function

Function Create_Feeder(ByVal length As Integer, ByVal nofeeder As Integer, ByVal feederlinecode As String)

    DSSText.Command = "New Line.Feeder" & nofeeder & ".1 Bus1=Main_Busbar Bus2=" & nofeeder & "_1 Length=1 units=m Linecode=" & feederlinecode
    For i = 1 To length - 1
        
        DSSText.Command = "New Line.Feeder" & nofeeder & "." & (i + 1) & " Bus1=" & nofeeder & "_" & i & " Bus2=" & nofeeder & "_" & (i + 1) & " Length=1 units=m Linecode=" & feederlinecode

    Next

End Function

Function Create_Lateral(ByVal length As Integer, ByVal nofeeder As Integer, ByVal nolateral As Integer, ByVal location As Integer, ByVal laterallinecode As String)


    DSSText.Command = "New Line.Lateral" & nofeeder & "." & nolateral & ".1 Bus1=" & nofeeder & "_" & location & " Bus2=" & nofeeder & "_" & nolateral & "_1 Length=1 units=m Linecode=" & laterallinecode
    For i = 1 To length - 1

        DSSText.Command = "New Line.Lateral" & nofeeder & "." & nolateral & "." & i + 1 & " Bus1=" & nofeeder & "_" & nolateral & "_" & i & " Bus2=" & nofeeder & "_" & nolateral & "_" & i + 1 & " Length=1 units=m Linecode=" & laterallinecode
        
    Next
End Function

Function Create_Consumer(ByVal nofeeder As Integer, ByVal nolateral As Integer, ByVal length As Integer, ByVal phase As Integer)

    

End Function



