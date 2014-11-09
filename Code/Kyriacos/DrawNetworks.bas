Attribute VB_Name = "DrawNetworks"
Sub DrawRural()
'
' DrawRural Macro

Dim customer As Integer
Dim workingsheet As Worksheet
Set workingsheet = Sheets("Network")
customer = 0

Sheets("Network").Activate
Cells.Select
Selection.Delete Shift:=xlUp

For y = 1 To 4
For i = 1 To 250
    workingsheet.Shapes.AddConnector(msoConnectorStraight, i * 5 - 5, 50 + (y * 500 - 500), i * 5, _
        50 + (y * 500 - 500)).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3.5
    End With
Next

For i = 1 To 196

    If i Mod 49 = 1 Then customer = customer + 1
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250, 50 + (y * 500 - 500) + (i * 2) - 2, 250, _
        50 + (y * 500 - 500) + (i * 2)).Select
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

For i = 1 To 196

    If i Mod 18 = 1 Then customer = customer + 1
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 150 * 5, 50 + (y * 500 - 500) + (i * 2) - 2, 150 * 5, _
        50 + (y * 500 - 500) + (i * 2)).Select
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

For i = 1 To 196

    If i Mod 22 = 1 Then customer = customer + 1
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250 * 5, 50 + (y * 500 - 500) + (i * 2) - 2, 250 * 5, _
        50 + (y * 500 - 500) + (i * 2)).Select
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


For i = 1 To 196

    If i Mod 22 = 1 Then customer = customer + 1
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250 * 5 + 40, 50 + (y * 500 - 500) + (i * 2) - 2, 250 * 5 + 40, _
        50 + (y * 500 - 500) + (i * 2)).Select
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
Call NetworkLabels

End Sub
Sub DrawSemiUrban()
'
' DrawRural Macro

Dim customer As Integer
Dim workingsheet As Worksheet
Set workingsheet = Sheets("Network")
customer = 0

Sheets("Network").Activate
Cells.Select
Selection.Delete Shift:=xlUp

For y = 1 To 4
For i = 1 To 250
    workingsheet.Shapes.AddConnector(msoConnectorStraight, i * 5 - 5, 50 + (y * 500 - 500), i * 5, _
        50 + (y * 500 - 500)).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3.5
    End With
Next

For i = 1 To 196

    If i Mod 17 = 1 Then customer = customer + 1 '12
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250, 50 + (y * 500 - 500) + (i * 2) - 2, 250, _
        50 + (y * 500 - 500) + (i * 2)).Select
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

For i = 1 To 195

    If i Mod 5 = 1 Then customer = customer + 1 '51
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 150 * 5, 50 + (y * 500 - 500) + (i * 2) - 2, 150 * 5, _
        50 + (y * 500 - 500) + (i * 2)).Select
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

For i = 1 To 196

    If i Mod 6 = 1 Then customer = customer + 1 '84
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250 * 5, 50 + (y * 500 - 500) + (i * 2) - 2, 250 * 5, _
        50 + (y * 500 - 500) + (i * 2)).Select
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


For i = 1 To 196

    If i Mod 6 = 1 Then customer = customer + 1 '117
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250 * 5 + 40, 50 + (y * 500 - 500) + (i * 2) - 2, 250 * 5 + 40, _
        50 + (y * 500 - 500) + (i * 2)).Select
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
Call NetworkLabels

End Sub

Sub DrawUrban()
'
' DrawRural Macro


Dim customer As Integer
Dim workingsheet As Worksheet
Set workingsheet = Sheets("Network")
customer = 0

Sheets("Network").Activate
Cells.Select
Selection.Delete Shift:=xlUp

For y = 1 To 4
For i = 1 To 250
    workingsheet.Shapes.AddConnector(msoConnectorStraight, i * 5 - 5, 50 + (y * 500 - 500), i * 5, _
        50 + (y * 500 - 500)).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .Weight = 3.5
    End With
Next

For i = 1 To 196

    If i Mod 12 = 1 Then customer = customer + 1
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250, 50 + (y * 500 - 500) + (i * 2) - 2, 250, _
        50 + (y * 500 - 500) + (i * 2)).Select
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

For i = 1 To 196 * 4

    If i Mod 15 = 1 Then customer = customer + 1
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 150 * 5, 50 + (y * 500 - 500) + (i * 0.5) - 0.5, 150 * 5, _
        50 + (y * 500 - 500) + (i * 0.5)).Select
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

For i = 1 To 196 * 2

    If i Mod 9 = 1 Then customer = customer + 1
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250 * 5, 50 + (y * 500 - 500) + (i * 1) - 1, 250 * 5, _
        50 + (y * 500 - 500) + (i * 1)).Select
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


For i = 1 To 196 * 2

    If i Mod 9 = 1 Then customer = customer + 1
        
    workingsheet.Shapes.AddConnector(msoConnectorStraight, 250 * 5 + 40, 50 + (y * 500 - 500) + (i * 1) - 1, 250 * 5 + 40, _
        50 + (y * 500 - 500) + (i * 1)).Select
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
Call NetworkLabels

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

