VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseNetwork 
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   OleObjectBlob   =   "ChooseNetwork.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChooseNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tday As Integer
Public finished As Boolean

Private Sub CHP_info_Click()
    HPinfo.Show
End Sub

Private Sub CHPPeneScroll_Change()
    CHPPeneText.Value = CHPPeneScroll.Value
End Sub

Private Sub CHPPeneText_Change()
    If CHPPeneText.Value <> "" Then CHPPeneScroll.Value = CHPPeneText.Value
    If Int(HPPeneText.Value) + Int(CHPPeneText.Value) > 100 Then
        HPPeneText.Value = 100 - CHPPeneText.Value
    End If
End Sub

Private Sub ClearnessScroll_Change()
    ClearnessText.Value = ClearnessScroll.Value
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
    
    If TdayVal.Value = "" Then
            MsgBox "Please select a type of day "
            Exit Sub
    End If
    
    If MonthVal.Value > 12 Or MonthVal.Value < 1 Then
        MsgBox "Please input a correct month"
        Exit Sub
    End If
    
    If TdayVal.Value <> "wd" And TdayVal.Value <> "we" Then
        MsgBox "Please input a correct type of day"
        Exit Sub
    End If
    
    If ChooseNetwork.PVPeneScroll.Value <> 0 Or ChooseNetwork.HPPeneScroll.Value <> 0 Or ChooseNetwork.CHPPeneScroll.Value <> 0 Then
        If SelectLocation.Value = "" Then
            MsgBox "Please select a location"
            Exit Sub
        End If
    End If
    
    If TdayVal.Value = "wd" Then Tday = 1 Else Tday = 2
    ChooseNetwork.Hide
    finished = True
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


Private Sub EVPeneScroll_Change()

    EVPeneText.Value = EVPeneScroll.Value

End Sub


Private Sub EVPeneText_Change()

   If EVPeneText.Value <> "" Then EVPeneScroll.Value = EVPeneText.Value
    
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub HP_info_Click()

    HPinfo.Show

End Sub
Private Sub HPPeneScroll_Change()

    HPPeneText.Value = HPPeneScroll.Value

End Sub


Private Sub HPPeneText_Change()

    If HPPeneText.Value <> "" Then HPPeneScroll.Value = HPPeneText.Value
    If Int(HPPeneText.Value) + Int(CHPPeneText.Value) > 100 Then
        CHPPeneText.Value = 100 - HPPeneText.Value
    End If
    
End Sub
Private Sub Label21_Click()

    SelectLocationForm.Show

End Sub

Private Sub Label22_Click()

    NetworkSpecifications.Show
    
End Sub

Private Sub PV_info_Click()

    PVinfo.Show

End Sub

Private Sub PVPeneScroll_Change()

    PVPeneText.Value = PVPeneScroll.Value
    
End Sub

Private Sub PVPeneText_Change()

    If PVPeneText.Value <> "" Then PVPeneScroll.Value = PVPeneText.Value
        
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
