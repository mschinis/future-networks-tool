Attribute VB_Name = "DrawNetworks"
Sub DrawBasicNetwork()
'

Dim customer As Integer
Dim WorkingSheet As Worksheet
Set WorkingSheet = Sheets("Network")


Sheets("Network").Activate
Cells.Select
Selection.Delete Shift:=xlUp

For y = 1 To 4

'Draw Feeder
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 0, 50 + (y * 500 - 500), 250 * 5, _
        50 + (y * 500 - 500)).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3.5
    End With

    'Draw Lateral 1
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 250, 50 + (y * 500 - 500), 250, _
        50 + (y * 500 - 500) + 450).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With
        
    'Draw Lateral 2
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 750, 50 + (y * 500 - 500), 150 * 5, _
        50 + (y * 500 - 500) + 450).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    'Draw Lateral 3
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 1250, 50 + (y * 500 - 500), 250 * 5, _
        50 + (y * 500 - 500) + 450).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3
    End With

    'Draw Lateral 4
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 1250 + 100, 50 + (y * 500 - 500), 250 * 5 + 100, _
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

Dim WorkingSheet As Worksheet
Set WorkingSheet = Sheets("Network")
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
        
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 250, (100 + (y * 500 - 500) + (400 * i / 12)), 250 + var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 750, (100 + (y * 500 - 500) + (400 * i / 39)), 750 + var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 1250, (100 + (y * 500 - 500) + (400 * i / 33)), 1250 - var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 1250 + 100, (100 + (y * 500 - 500) + (400 * i / 33)), 1250 + 100 + var, _
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
'Call NetworkLabels

End Sub

Sub DrawUrban()
'


Dim customer As Integer
Dim WorkingSheet As Worksheet
Set WorkingSheet = Sheets("Network")
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
        
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 250, (100 + (y * 500 - 500) + (400 * i / 17)), 250 + var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 750, (100 + (y * 500 - 500) + (400 * i / 53)), 750 + var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 1250, (100 + (y * 500 - 500) + (400 * i / 44)), 1250 - var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 1250 + 100, (100 + (y * 500 - 500) + (400 * i / 44)), 1250 + 100 + var, _
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
'Call NetworkLabels

End Sub

Sub DrawRural()
'
' DrawRural Macro

Dim customer As Integer
Dim WorkingSheet As Worksheet
Set WorkingSheet = Sheets("Network")
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
        
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 250, (100 + (y * 500 - 500) + (400 * i / 4)), 250 + var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 750, (100 + (y * 500 - 500) + (400 * i / 11)), 750 + var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 1250, (100 + (y * 500 - 500) + (400 * i / 9)), 1250 + var, _
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
    WorkingSheet.Shapes.AddConnector(msoConnectorStraight, 1250 + 100, (100 + (y * 500 - 500) + (400 * i / 9)), 1250 + 100 + var, _
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
'Call NetworkLabels

End Sub

Sub NetworkLabels()
'
' NetworkLabels Macro
'

'
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Feeder 1"
    Range("A3").Select
    Selection.Font.Bold = True
    With Selection.Font
        .name = "Arial"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("C9").Select

    Range("A42").Select

    Range("A3").Select
    Selection.Copy

    Range("A42").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Feeder 2"
    Range("A43").Select

    Range("A81").Select

    Range("A42").Select
    Selection.Copy

    Range("A81").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Feeder 3"
    Range("A82").Select

    Range("A120").Select

    Range("A81").Select
    Selection.Copy

    Range("A120").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Feeder 4"
    Range("A121").Select

    Range("F37").Select

    ActiveCell.FormulaR1C1 = "Lateral 1"
    Range("F37").Select
    Selection.Font.Bold = True
    With Selection.Font
        .name = "Arial"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("P37").Select

    ActiveCell.FormulaR1C1 = ""
    Range("F37").Select
    Selection.Copy
    Range("P37").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Lateral 2"
    Range("P37").Select
    Selection.Copy
    Range("Z37").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Lateral 3"
    Range("Z37").Select
    Selection.Copy
    Range("AB35").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Lateral 4"
    Range("AB36").Select

    Range("F37").Select
    Selection.Copy

    Range("F76").Select
    ActiveSheet.Paste

    Range("P37").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("P76").Select
    ActiveSheet.Paste

    Range("Z37").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("Z76").Select
    ActiveSheet.Paste

    Range("AB35").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("AB75").Select

    Range("AB74").Select

    ActiveSheet.Paste

    Range("F76").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("F115").Select
    ActiveSheet.Paste

    Range("P76").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("P115").Select
    ActiveSheet.Paste

    Range("Z76").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("Z115").Select
    ActiveSheet.Paste

    Range("AB74").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("AB114").Select
    ActiveSheet.Paste

    Range("F115").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("F154").Select
    ActiveSheet.Paste
 
    Range("P115").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("P154").Select
    ActiveSheet.Paste

    Range("Z115").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("Z154").Select

    ActiveSheet.Paste
    Range("AB114").Select
    Application.CutCopyMode = False
    Selection.Copy

    Range("AB153").Select
    ActiveSheet.Paste

    'ActiveWindow.Zoom = 70
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.SmallScroll Down:=-200
End Sub

Sub CurrentOverload()
'
' CurrentOverload Macro
    
    For i = 1 To 4
        For y = 1 To 5
        If Start.CurrentFlags(i, y) = 1 Then
            Sheets("Network").Shapes("Feeder1Lateral1").Visible = True
        Else
'    Dim LateralNo(1 To 5) As Double
'    Dim FeederNo(1 To 4) As Double
'
'    FeederNo(1) = 64
'    FeederNo(2) = 585
'    FeederNo(3) = 1104
'    FeederNo(4) = 1625
'
'    LateralNo(1) = 6
'    LateralNo(2) = 259
'    LateralNo(3) = 759
'    LateralNo(4) = 1181
'    LateralNo(5) = 1298
'
'
'    For i = 1 To 4
'        For y = 1 To 5
'            If Start.CurrentFlags(i, y) = 1 Then
'
'                ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, LateralNo(y), FeederNo(i), 65.25, 37.5).Select
'                Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "CURRENT EXCEEDED"
'                With Selection.ShapeRange.Fill
'                    .Visible = msoTrue
'                    .ForeColor.RGB = RGB(255, 0, 0)
'                    .Transparency = 0
'                    .Solid
'                End With
'
'                With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 16). _
'                    ParagraphFormat
'                    .FirstLineIndent = 0
'                    .Alignment = msoAlignCenter
'                End With
'                With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 16).Font
'                    .NameComplexScript = "+mn-cs"
'                    .NameFarEast = "+mn-ea"
'                    .Fill.Visible = msoTrue
'                    .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
'                    .Fill.ForeColor.TintAndShade = 0
'                    .Fill.ForeColor.Brightness = 0
'                    .Fill.Transparency = 0
'                    .Fill.Solid
'                    .Size = 11
'                    .name = "+mn-lt"
'                End With
'
'            End If
'        Next
'    Next
'
'
End Sub


