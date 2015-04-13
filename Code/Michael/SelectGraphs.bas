Attribute VB_Name = "SelectGraphs"
Public Sub setupSelectGraphsForm()
    Dim counter, feeders, laterals As Integer

    Dim settingsForm As Object
    Dim NewFrame As MSForms.Frame
    Dim NewButton As MSForms.CommandButton
    Dim NewListBox As MSForms.ListBox
    Dim X As Integer
    Dim Line As Integer
    
    'This is to stop screen flashing while creating form
    'Application.VBE.MainWindow.Visible = False
    
    counter = 1
    feeders = Sheets("LastSimulationData").Range("B3").Value
    laterals = Sheets("LastSimulationData").Range("B4").Value
    
    'Create the User Form
    With [Form_Site Revenue ]
        .Caption = "Yolo!"
        .Width = 110 * feeders
        .Height = 270
    End With
    
    SelectGraphsForm.Show
    Exit Sub
    
    
    
    
    'Create ListBox
    For i = 1 To feeders
        Set NewListBox = settingsForm.designer.Controls.Add("Forms.listbox.1")
        With NewListBox
            .name = "fdr_1"
            .Top = 10
            .Left = i * 3 + (i - 1) * 105
            .Width = 105
            .Height = 180
            .Font.Size = 8
            .Font.name = "Tahoma"
            .BorderStyle = fmBorderStyleOpaque
            .SpecialEffect = fmSpecialEffectSunken
        End With
    Next i
    ' Generate the filling of the listBoxes
    
    settingsForm.codemodule.insertlines 1, "Private Sub UserForm_Initialize()"
    For i = 1 To feeders
        myForm.codemodule.insertlines 2, "  me.frm_" & i & ".AddItem ""Lateral 1 Start Voltage"" "
    Next i
    'add code for listBox
    
    myForm.codemodule.insertlines 2, "   me.lst_1.addItem ""Data 1"" "
    myForm.codemodule.insertlines 3, "   me.lst_1.addItem ""Data 2"" "
    myForm.codemodule.insertlines 4, "   me.lst_1.addItem ""Data 3"" "
    myForm.codemodule.insertlines 5, "End Sub"
    
    
    'Create CommandButton Create
    Set NewButton = myForm.designer.Controls.Add("Forms.commandbutton.1")
    With NewButton
        .name = "cmd_1"
        .Caption = "clickMe"
        .Accelerator = "M"
        .Top = 10
        .Left = 200
        .Width = 66
        .Height = 20
        .Font.Size = 8
        .Font.name = "Tahoma"
        .BackStyle = fmBackStyleOpaque
    End With
    
   
    
    'add code for Comand Button
    myForm.codemodule.insertlines 6, "Private Sub cmd_1_Click()"
    myForm.codemodule.insertlines 7, "   If me.lst_1.text <>"""" Then"
    myForm.codemodule.insertlines 8, "      msgbox (""You selected item: "" & me.lst_1.text )"
    myForm.codemodule.insertlines 9, "   End If"
    myForm.codemodule.insertlines 10, "End Sub"
    'Show the form
    VBA.UserForms.Add(myForm.name).Show
    
    'Delete the form (Optional)
    'ThisWorkbook.VBProject.VBComponents.Remove myForm

End Sub





























Public Sub setupSelectGraphsForm2()
    Dim counter, feeders, laterals As Integer

    'SelectGraphsForm.Show
    Dim settingsForm As Object
    Dim NewFrame As MSForms.Frame
    Dim NewButton As MSForms.CommandButton
    Dim NewListBox As MSForms.ListBox
    Dim X As Integer
    Dim Line As Integer
    
    'This is to stop screen flashing while creating form
    Application.VBE.MainWindow.Visible = False
    
    Set counter = 1
    Set feeders = Sheets("LastSimulationData").Range("B3").Value
    Set laterals = Sheets("LastSimulationData").Range("B4").Value
    
    Set myForm = ThisWorkbook.VBProject.VBComponents.Add(3)
    'Create the User Form
    With settingsForm
        .Properties("Caption") = "Select Graphs"
        .Properties("Width") = 110 * feeders
        .Properties("Height") = 270
    End With
    
    'Create ListBox
    For i = 1 To feeders
        Set NewListBox = settingsForm.designer.Controls.Add("Forms.listbox.1")
        With NewListBox
            .name = "fdr_1"
            .Top = 10
            .Left = i * 3 + (i - 1) * 105
            .Width = 105
            .Height = 180
            .Font.Size = 8
            .Font.name = "Tahoma"
            .BorderStyle = fmBorderStyleOpaque
            .SpecialEffect = fmSpecialEffectSunken
        End With
    Next i
    ' Generate the filling of the listBoxes
    
    settingsForm.codemodule.insertlines 1, "Private Sub UserForm_Initialize()"
    For i = 1 To feeders
        myForm.codemodule.insertlines 2, "  me.frm_" & i & ".AddItem ""Lateral 1 Start Voltage"" "
        .AddItem "Lateral 2 Start Voltage"
        .AddItem "Lateral 3 Start Voltage"
        .AddItem "Lateral 4 Start Voltage"
        
        .AddItem "Lateral 1 End Voltage"
        .AddItem "Lateral 2 End Voltage"
        .AddItem "Lateral 3 End Voltage"
        .AddItem "Lateral 4 End Voltage"
        
        .AddItem "Feeder 1 Start Current"
        
        .AddItem "Lateral 1 Start Current"
        .AddItem "Lateral 2 Start Current"
        .AddItem "Lateral 3 Start Current"
        .AddItem "Lateral 4 Start Current"
    Next i
    'add code for listBox
    
    myForm.codemodule.insertlines 2, "   me.lst_1.addItem ""Data 1"" "
    myForm.codemodule.insertlines 3, "   me.lst_1.addItem ""Data 2"" "
    myForm.codemodule.insertlines 4, "   me.lst_1.addItem ""Data 3"" "
    myForm.codemodule.insertlines 5, "End Sub"
    
    
    'Create CommandButton Create
    Set NewButton = myForm.designer.Controls.Add("Forms.commandbutton.1")
    With NewButton
        .name = "cmd_1"
        .Caption = "clickMe"
        .Accelerator = "M"
        .Top = 10
        .Left = 200
        .Width = 66
        .Height = 20
        .Font.Size = 8
        .Font.name = "Tahoma"
        .BackStyle = fmBackStyleOpaque
    End With
    
   
    
    'add code for Comand Button
    myForm.codemodule.insertlines 6, "Private Sub cmd_1_Click()"
    myForm.codemodule.insertlines 7, "   If me.lst_1.text <>"""" Then"
    myForm.codemodule.insertlines 8, "      msgbox (""You selected item: "" & me.lst_1.text )"
    myForm.codemodule.insertlines 9, "   End If"
    myForm.codemodule.insertlines 10, "End Sub"
    'Show the form
    VBA.UserForms.Add(myForm.name).Show
    
    'Delete the form (Optional)
    'ThisWorkbook.VBProject.VBComponents.Remove myForm

End Sub

