VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseNetwork 
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4230
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

Private Sub CHPEnable_Click()

    If CHPEnable = False Then
        Label16.Visible = False
        CHPPeneText.Visible = False
        CHPPeneScroll.Visible = False

    Else
        Label16.Visible = True
        CHPPeneText.Visible = True
        CHPPeneScroll.Visible = True

    End If
End Sub

Private Sub CHPPeneScroll_Change()

    CHPPeneText.Value = CHPPeneScroll.Value

End Sub

Private Sub CHPPeneText_Change()

    If CHPPeneText.Value <> "" Then CHPPeneScroll.Value = CHPPeneText.Value

End Sub

Private Sub ClearnessScroll_Change()

    ClearnessText.Value = ClearnessScroll.Value

End Sub

Private Sub CommandButton1_Click()

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
    
    If PVEnable.Value = True Or HPEnable.Value = True Or CHPEnable.Value = True Then
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



Private Sub EVEnable_Click()

    If EVEnable = True Then
        Label6.Visible = True
        EVPeneText.Visible = True
        EVPeneScroll.Visible = True
        
    Else
        Label6.Visible = False
        EVPeneText.Visible = False
        EVPeneScroll.Visible = False
    End If

End Sub

Private Sub EVPeneScroll_Change()

    EVPeneText.Value = EVPeneScroll.Value

End Sub


Private Sub EVPeneText_Change()

   If EVPeneText.Value <> "" Then EVPeneScroll.Value = EVPeneText.Value
    
End Sub

Private Sub HPEnable_Click()

    If HPEnable = False Then
        Label12.Visible = False

        HPPeneText.Visible = False
        HPPeneScroll.Visible = False


    Else
        Label12.Visible = True

        HPPeneText.Visible = True
        HPPeneScroll.Visible = True

    End If
End Sub

Private Sub HPPeneScroll_Change()

    HPPeneText.Value = HPPeneScroll.Value

End Sub


Private Sub HPPeneText_Change()

    If HPPeneText.Value <> "" Then HPPeneScroll.Value = HPPeneText.Value
    
End Sub

Private Sub Label21_Click()

    SelectLocationForm.Show

End Sub

Private Sub Label22_Click()

    NetworkSpecifications.Show
    
End Sub



Private Sub PVEnable_Click()
    
    If PVEnable = False Then
        Label8.Visible = False

        PVPeneText.Visible = False
        PVPeneScroll.Visible = False

        ClearnessText.Visible = False
        ClearnessScroll.Visible = False
        Label11.Visible = False
    Else
        Label8.Visible = True

        PVPeneText.Visible = True
        PVPeneScroll.Visible = True

        ClearnessText.Visible = True
        ClearnessScroll.Visible = True
        Label11.Visible = True
    End If

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

    'EVPage
    Label6.Visible = False
    EVPeneText.Visible = False
    EVPeneScroll.Visible = False
    EVPeneText.Value = 0

    'PVPage
    Label8.Visible = False

    PVPeneText.Visible = False
    PVPeneScroll.Visible = False
    PVPeneText.Value = 0

    ClearnessText.Value = 1
    ClearnessText.Visible = False
    ClearnessScroll.Visible = False
    Label11.Visible = False
    
    'HPPage
    Label12.Visible = False

    HPPeneText.Visible = False
    HPPeneScroll.Visible = False

    HPPeneText.Value = 0
    
    'CHPPage
    Label16.Visible = False

    CHPPeneText.Visible = False
    CHPPeneScroll.Visible = False

    CHPPeneText.Value = 0
    
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
