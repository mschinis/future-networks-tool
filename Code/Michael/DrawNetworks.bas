Attribute VB_Name = "DrawNetworks"
Sub DrawBasicNetwork()
'

Dim customer As Integer
Dim workingsheet As Worksheet
Set workingsheet = Sheets("Network")
Sheets("Network").Activate

Dim shp As Shape
For Each shp In ActiveSheet.Shapes
    If shp.Type = 1 Then
        shp.Delete
    End If
Next shp


Sheets("Network").Activate
'Cells.Select
'Selection.Delete Shift:=xlUp

For y = 1 To 4

'Draw Feeder
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 0, 50 + (y * 500 - 500), 1350, _
        50 + (y * 500 - 500)).Select
    With Selection.ShapeRange.Line
        .Weight = 4
        .Visible = msoTrue
    End With

    'Draw Lateral 1
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250, 50 + (y * 500 - 500), 250, _
        50 + (y * 500 - 500) + 450).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With
        
    'Draw Lateral 2
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 750, 50 + (y * 500 - 500), 150 * 5, _
        50 + (y * 500 - 500) + 450).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    'Draw Lateral 3
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 1250, 50 + (y * 500 - 500), 250 * 5, _
        50 + (y * 500 - 500) + 450).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    'Draw Lateral 4
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 1250 + 100, 50 + (y * 500 - 500), 250 * 5 + 100, _
        50 + (y * 500 - 500) + 450).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

Next


End Sub
Sub DrawSemiUrban()
'
' DrawRural Macro

Dim customer As Integer
Dim var As Integer

Dim workingsheet As Worksheet
Set workingsheet = Sheets("Network")
customer = 0

Call DrawBasicNetwork


For y = 1 To 4

For i = 1 To 12
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250, (100 + (y * 500 - 500) + (400 * i / 12)), 250 + var, _
        (100 + (y * 500 - 500) + (400 * i / 12))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 39
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 750, (100 + (y * 500 - 500) + (400 * i / 39)), 750 + var, _
        (100 + (y * 500 - 500) + (400 * i / 39))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 33
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 1250, (100 + (y * 500 - 500) + (400 * i / 33)), 1250 + var, _
        (100 + (y * 500 - 500) + (400 * i / 33))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 33
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 1250 + 100, (100 + (y * 500 - 500) + (400 * i / 33)), 1250 + 100 + var, _
        (100 + (y * 500 - 500) + (400 * i / 33))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

Next

    ActiveSheet.Shapes.Range(Array("Group 1")).Select
    Selection.ShapeRange.ZOrder msoBringToFront
End Sub

Sub DrawUrban()

Dim customer As Integer
Dim workingsheet As Worksheet
Set workingsheet = Sheets("Network")
customer = 0

Call DrawBasicNetwork

For y = 1 To 4

For i = 1 To 17
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250, (100 + (y * 500 - 500) + (400 * i / 17)), 250 + var, _
        (100 + (y * 500 - 500) + (400 * i / 17))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 53
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 750, (100 + (y * 500 - 500) + (400 * i / 53)), 750 + var, _
        (100 + (y * 500 - 500) + (400 * i / 53))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 44
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 1250, (100 + (y * 500 - 500) + (400 * i / 44)), 1250 + var, _
        (100 + (y * 500 - 500) + (400 * i / 44))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 44
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 1250 + 100, (100 + (y * 500 - 500) + (400 * i / 44)), 1250 + 100 + var, _
        (100 + (y * 500 - 500) + (400 * i / 44))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

Next

    ActiveSheet.Shapes.Range(Array("Group 1")).Select
    Selection.ShapeRange.ZOrder msoBringToFront
End Sub

Sub DrawRural()
'
' DrawRural Macro

Dim customer As Integer
Dim workingsheet As Worksheet
Set workingsheet = Sheets("Network")
customer = 0

Call DrawBasicNetwork


For y = 1 To 4

For i = 1 To 4
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250, (100 + (y * 500 - 500) + (400 * i / 4)), 250 + var, _
        (100 + (y * 500 - 500) + (400 * i / 4))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 11
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 750, (100 + (y * 500 - 500) + (400 * i / 11)), 750 + var, _
        (100 + (y * 500 - 500) + (400 * i / 11))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 9
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 1250, (100 + (y * 500 - 500) + (400 * i / 9)), 1250 + var, _
        (100 + (y * 500 - 500) + (400 * i / 9))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

For i = 1 To 9
    customer = customer + 1
    If i Mod 2 = 1 Then
        var = 30
    Else
        var = -30
    End If
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 1250 + 100, (100 + (y * 500 - 500) + (400 * i / 9)), 1250 + 100 + var, _
        (100 + (y * 500 - 500) + (400 * i / 9))).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    If Start.CustomerVoltageLimit(customer) = 1 Then

        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
    Else
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With
    End If
Next

Next
    ActiveSheet.Shapes.Range(Array("Group 1")).Select
    Selection.ShapeRange.ZOrder msoBringToFront

End Sub



Sub CurrentOverload()


   Dim CurrentFlags(1 To 4, 1 To 5) As Boolean
    
    If Sheets("FeederCurrentRollingAverages").Range("C1390").Value > 1 Or Sheets("FeederCurrentRollingAverages").Range("C1389").Value < -1 Then CurrentFlags(1, 1) = True
    If Sheets("CurrentRollingAverages").Range("E1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("E1391").Value < -1 Then CurrentFlags(1, 2) = True
    If Sheets("CurrentRollingAverages").Range("F1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("F1391").Value < -1 Then CurrentFlags(1, 3) = True
    If Sheets("CurrentRollingAverages").Range("I1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("I1391").Value < -1 Then CurrentFlags(1, 4) = True
    If Sheets("CurrentRollingAverages").Range("L1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("L1391").Value < -1 Then CurrentFlags(1, 5) = True
    
    If Sheets("FeederCurrentRollingAverages").Range("F1390").Value > 1 Or Sheets("FeederCurrentRollingAverages").Range("F1389").Value < -1 Then CurrentFlags(2, 1) = True
    If Sheets("CurrentRollingAverages").Range("O1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("O1391").Value < -1 Then CurrentFlags(2, 2) = True
    If Sheets("CurrentRollingAverages").Range("R1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("R1391").Value < -1 Then CurrentFlags(2, 3) = True
    If Sheets("CurrentRollingAverages").Range("U1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("U1391").Value < -1 Then CurrentFlags(2, 4) = True
    If Sheets("CurrentRollingAverages").Range("X1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("X1391").Value < -1 Then CurrentFlags(2, 5) = True
    
    If Sheets("FeederCurrentRollingAverages").Range("I1390").Value > 1 Or Sheets("FeederCurrentRollingAverages").Range("I1389").Value < -1 Then CurrentFlags(3, 1) = True
    If Sheets("CurrentRollingAverages").Range("AA1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("AA1391").Value < -1 Then CurrentFlags(3, 2) = True
    If Sheets("CurrentRollingAverages").Range("AD1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("AD1391").Value < -1 Then CurrentFlags(3, 3) = True
    If Sheets("CurrentRollingAverages").Range("AG1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("AG1391").Value < -1 Then CurrentFlags(3, 4) = True
    If Sheets("CurrentRollingAverages").Range("AJ1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("AJ1391").Value < -1 Then CurrentFlags(3, 5) = True
    
    If Sheets("FeederCurrentRollingAverages").Range("L1390").Value > 1 Or Sheets("FeederCurrentRollingAverages").Range("L1389").Value < -1 Then CurrentFlags(4, 1) = True
    If Sheets("CurrentRollingAverages").Range("AM1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("AM1391").Value < -1 Then CurrentFlags(4, 2) = True
    If Sheets("CurrentRollingAverages").Range("AP1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("AP1391").Value < -1 Then CurrentFlags(4, 3) = True
    If Sheets("CurrentRollingAverages").Range("AS1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("AS1391").Value < -1 Then CurrentFlags(4, 4) = True
    If Sheets("CurrentRollingAverages").Range("AV1392").Value > 1 Or Sheets("CurrentRollingAverages").Range("AV1391").Value < -1 Then CurrentFlags(4, 5) = True
    
    For i = 1 To 4
        For y = 1 To 5
        If CurrentFlags(i, y) = True Then
            Sheets("Network").Shapes("Feeder" & i & "Lateral" & y - 1).Visible = True
        Else
            Sheets("Network").Shapes("Feeder" & i & "Lateral" & y - 1).Visible = False
        End If
        Next
    Next

End Sub


