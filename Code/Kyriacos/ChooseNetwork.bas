VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseNetwork 
   ClientHeight    =   12810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   OleObjectBlob   =   "ChooseNetwork.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChooseNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tday As Integer

Private Sub CHP_info_Click()
    HPinfo.Show
End Sub

Private Sub CHPPeneScroll_Change()
    CHPPeneText.Value = CHPPeneScroll.Value
End Sub

Private Sub CHPPeneScroll2_Change()
    
    CHPPeneText2.Value = CHPPeneScroll2.Value
    If FeedersBox.Value <> "" And LateralsBox.Value <> "" And LateralsBox.Value <> "All Laterals" Then
        Assign_Profiles.CHPPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) = Int(CHPPeneText2.Value) / 100
    ElseIf FeedersBox.Value <> "" And LateralsBox.Value = "All Laterals" Then
        For i = 1 To Assign_Profiles.NoLaterals
            Assign_Profiles.CHPPenetrationArray(Int(FeedersBox.Value), i) = Int(CHPPeneText2.Value) / 100
        Next
    Else
        CHPPeneText2.Value = 0
    End If
End Sub

Private Sub CHPPeneText_Change()
    If CHPPeneText.Value <> "" Then CHPPeneScroll.Value = CHPPeneText.Value
    If Int(HPPeneText.Value) + Int(CHPPeneText.Value) > 100 Then
        HPPeneText.Value = 100 - CHPPeneText.Value
    End If
End Sub

Private Sub CHPPeneText2_Change()
    
    If CHPPeneText2.Value <> "" Then CHPPeneScroll2.Value = CHPPeneText2.Value
    If CHPPeneText2.Value <> "" Then CHPPeneScroll2.Value = CHPPeneText2.Value
    If Int(HPPeneText2.Value) + Int(CHPPeneText2.Value) > 100 Then
        HPPeneText2.Value = 100 - CHPPeneText2.Value
    End If
End Sub

Private Sub ClearnessScroll_Change()
    ClearnessText.Value = ClearnessScroll.Value
End Sub

Private Sub CommandButton3_Click()
    
    MSG1 = MsgBox("Are you sure you want to reset all Lateral Specific penetration values?", vbYesNo, "Warning!")

    If MSG1 = vbYes Then
            SelectNetwork_Change
            PVPeneText2.Value = 0
            EVPeneText2.Value = 0
            HPPeneText2.Value = 0
            CHPPeneText2.Value = 0
        
    Else
        
    End If

    

End Sub

Private Sub ContinueBtn_Click()
    
    If SelectNetwork.Value = "" Then
            MsgBox "Please select a network"
            Exit Sub
    End If
    If MonthVal.Value = "" Then
            MsgBox "Please select a month"
            Exit Sub
    End If
    
    If MonthVal.Value > 12 Or MonthVal.Value < 1 Then
        MsgBox "Please input a correct month"
        Exit Sub
    End If
    
    If ChooseNetwork.PVPeneScroll.Value <> 0 Or ChooseNetwork.HPPeneScroll.Value <> 0 Or ChooseNetwork.CHPPeneScroll.Value <> 0 Then
        If SelectLocation.Value = "" Then
            MsgBox "Please select a location"
            Exit Sub
        End If
    End If
    If TdayOptionWD.Value = True Then Tday = 1 Else Tday = 2
    
    ChooseNetwork.Hide
    Start.finished = True
    Preset_Network

End Sub



Private Sub CommandButton2_Click()
    
    If SelectNetwork.Value = "" Or MonthVal.Value = "" Or MonthVal.Value > 12 Or MonthVal.Value < 1 Then
        msg2 = MsgBox("Please select a network and/or input a valid month before accessing the Advanced Settings", vbOKOnly, "Warning!")
    Else
        AdvancedProperties.Show
    End If
    

End Sub

Private Sub EV_Info_Click()
    EVinfo.Show
End Sub

Private Sub EVFeeder_Change()

End Sub

Private Sub EVPeneScroll_Change()
    EVPeneText.Value = EVPeneScroll.Value
End Sub

Private Sub EVPeneScroll2_Change()
    EVPeneText2.Value = EVPeneScroll2.Value
    
    If FeedersBox.Value <> "" And LateralsBox.Value <> "" And LateralsBox.Value <> "All Laterals" Then
        Assign_Profiles.EVPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) = Int(EVPeneText2.Value) / 100
    ElseIf FeedersBox.Value <> "" And LateralsBox.Value = "All Laterals" Then
        For i = 1 To Assign_Profiles.NoLaterals
            Assign_Profiles.EVPenetrationArray(Int(FeedersBox.Value), i) = Int(EVPeneText2.Value) / 100
        Next
    Else
        EVPeneText2.Value = 0
    End If
    
End Sub

Private Sub EVPeneText_Change()
   If EVPeneText.Value <> "" Then EVPeneScroll.Value = EVPeneText.Value
End Sub


Private Sub EVPeneText2_Change()
   If EVPeneText2.Value <> "" Then EVPeneScroll2.Value = EVPeneText2.Value
End Sub

Private Sub FeedersBox_Change()

    If FeedersBox.Value <> "" And LateralsBox.Value <> "" And LateralsBox.Value <> "All Laterals" Then
        PVPeneText2.Value = Assign_Profiles.PVPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) * 100
        EVPeneText2.Value = Assign_Profiles.EVPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) * 100
        HPPeneText2.Value = Assign_Profiles.HPPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) * 100
        CHPPeneText2.Value = Assign_Profiles.CHPPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) * 100
    ElseIf FeedersBox.Value <> "" And LateralsBox.Value = "All Laterals" Then

        PVPeneText2.Value = Assign_Profiles.PVPenetrationArray(Int(FeedersBox.Value), 1) * 100
        EVPeneText2.Value = Assign_Profiles.EVPenetrationArray(Int(FeedersBox.Value), 1) * 100
        HPPeneText2.Value = Assign_Profiles.HPPenetrationArray(Int(FeedersBox.Value), 1) * 100
        CHPPeneText2.Value = Assign_Profiles.CHPPenetrationArray(Int(FeedersBox.Value), 1) * 100
    End If
    
End Sub

Private Sub Frame6_Click()

End Sub

Private Sub HP_info_Click()

    HPinfo.Show

End Sub
Private Sub HPPeneScroll_Change()

    HPPeneText.Value = HPPeneScroll.Value

End Sub


Private Sub HPPeneScroll2_Change()
    HPPeneText2.Value = HPPeneScroll2.Value
    
    If FeedersBox.Value <> "" And LateralsBox.Value <> "" And LateralsBox.Value <> "All Laterals" Then
        Assign_Profiles.HPPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) = Int(HPPeneText2.Value) / 100
    ElseIf FeedersBox.Value <> "" And LateralsBox.Value = "All Laterals" Then
        For i = 1 To Assign_Profiles.NoLaterals
            Assign_Profiles.HPPenetrationArray(Int(FeedersBox.Value), i) = Int(HPPeneText2.Value) / 100
        Next
    Else
        HPPeneText2.Value = 0
    End If
End Sub

Private Sub HPPeneText_Change()

    If HPPeneText.Value <> "" Then HPPeneScroll.Value = HPPeneText.Value
    If Int(HPPeneText.Value) + Int(CHPPeneText.Value) > 100 Then
        CHPPeneText.Value = 100 - HPPeneText.Value
    End If
    
End Sub

Private Sub HPPeneText2_Change()
    
    If HPPeneText2.Value <> "" Then HPPeneScroll2.Value = HPPeneText2.Value
    If HPPeneText2.Value <> "" Then HPPeneScroll2.Value = HPPeneText2.Value
    If Int(HPPeneText2.Value) + Int(CHPPeneText2.Value) > 100 Then
        CHPPeneText2.Value = 100 - HPPeneText2.Value
    End If
End Sub

Private Sub Label21_Click()

    SelectLocationForm.Show

End Sub

Private Sub Label22_Click()

    NetworkSpecifications.Show
    
End Sub

Private Sub LateralsBox_Change()
    
    If FeedersBox.Value <> "" And LateralsBox.Value <> "" And LateralsBox.Value <> "All Laterals" Then
        PVPeneText2.Value = Assign_Profiles.PVPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) * 100
        EVPeneText2.Value = Assign_Profiles.EVPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) * 100
        HPPeneText2.Value = Assign_Profiles.HPPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) * 100
        CHPPeneText2.Value = Assign_Profiles.CHPPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) * 100
    ElseIf FeedersBox.Value <> "" And LateralsBox.Value = "All Laterals" Then

        PVPeneText2.Value = Assign_Profiles.PVPenetrationArray(Int(FeedersBox.Value), 1) * 100
        EVPeneText2.Value = Assign_Profiles.EVPenetrationArray(Int(FeedersBox.Value), 1) * 100
        HPPeneText2.Value = Assign_Profiles.HPPenetrationArray(Int(FeedersBox.Value), 1) * 100
        CHPPeneText2.Value = Assign_Profiles.CHPPenetrationArray(Int(FeedersBox.Value), 1) * 100
    End If
End Sub



Private Sub PV_info_Click()

    PVinfo.Show

End Sub

Private Sub PVPeneScroll_Change()

    PVPeneText.Value = PVPeneScroll.Value
    
End Sub

Private Sub PVPeneScroll2_Change()
    PVPeneText2.Value = PVPeneScroll2.Value
    
    If FeedersBox.Value <> "" And LateralsBox.Value <> "" And LateralsBox.Value <> "All Laterals" Then
        Assign_Profiles.PVPenetrationArray(Int(FeedersBox.Value), Int(LateralsBox.Value)) = Int(PVPeneText2.Value) / 100
    ElseIf FeedersBox.Value <> "" And LateralsBox.Value = "All Laterals" Then
        For i = 1 To Assign_Profiles.NoLaterals
            Assign_Profiles.PVPenetrationArray(Int(FeedersBox.Value), i) = Int(PVPeneText2.Value) / 100
        Next
    Else
        PVPeneText2.Value = 0
    End If
End Sub

Private Sub PVPeneText_Change()

    If PVPeneText.Value <> "" Then PVPeneScroll.Value = PVPeneText.Value
        
End Sub

Private Sub PVPeneText2_Change()
    If PVPeneText2.Value <> "" Then PVPeneScroll2.Value = PVPeneText2.Value
End Sub

Private Sub SelectNetwork_Change()

    If SelectNetwork.Value = "Urban" Or SelectNetwork.Value = "SemiUrban" Or SelectNetwork.Value = "Rural" Then
        Assign_Profiles.NoLaterals = 4
        Assign_Profiles.NoFeeders = 4
        
        ReDim Assign_Profiles.PVPenetrationArray(1 To 4, 1 To 4)
        ReDim Assign_Profiles.EVPenetrationArray(1 To 4, 1 To 4)
        ReDim Assign_Profiles.HPPenetrationArray(1 To 4, 1 To 4)
        ReDim Assign_Profiles.CHPPenetrationArray(1 To 4, 1 To 4)
        
        Assign_Profiles.LateralSizes = PresetLateralSizes(SelectNetwork.Value)

        FeedersBox.Clear
        LateralsBox.Clear
        
        With FeedersBox
            For i = 1 To Assign_Profiles.NoFeeders
                .AddItem i
            Next
        End With
        
        With LateralsBox
            For i = 1 To Assign_Profiles.NoFeeders
                .AddItem i
            Next
            .AddItem "All Laterals"
        End With
        
    End If
    


End Sub

Public Sub UserForm_Initialize()
    Dim filename As String
        
    filename = Dir(ThisWorkbook.Path & "/Networks/", 16)
    filename = Dir()
    filename = Dir()
    Do While filename <> ""
        If filename <> "Custom" Then
            SelectNetwork.AddItem filename
        End If
    filename = Dir()
    Loop
    
    With SelectLocation
        .AddItem "Scotland"
        .AddItem "North East"
        .AddItem "North West"
        .AddItem "Yorkshire and Humber"
        .AddItem "East Midlands"
        .AddItem "West Midlands"
        .AddItem "East"
        .AddItem "Wales"
        .AddItem "London"
        .AddItem "South East"
        .AddItem "South West"
    End With
    
    
    With TransformerTap
        .AddItem "-5"
        .AddItem "-2.5"
        .AddItem "0"
        .AddItem "2.5"
        .AddItem "5"
    End With
    
    TransformerTap.Value = "0"


    
End Sub
