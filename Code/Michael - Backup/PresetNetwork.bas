Attribute VB_Name = "PresetNetwork"
Public Network As String
Public customers As Integer


Public Sub Preset_Network()

Dim stime As Single
    Dim EVPenetration, PVPenetration, HPPenetration, CHPPenetration As Double
    Dim location, Tmonth, Tday, clearness As Integer
    Network = ChooseNetwork.SelectNetwork.Value ' Select Network from Dropdown Menu
    Dim File_Location As String
    
    Assign_Profiles.CHPStopPoint = 0
    Assign_Profiles.HPStopPoint = 0
    
    File_Location = "Networks\" & Trim(Network) & "\" & Trim(Network)
    
    File_Exists_Check = miscMacros.File_Exists(File_Location & ".dss")
    If File_Exists_Check = False Then
        MsgBox ActiveWorkbook.Path & "\" & File_Location & ".dss file not found."
        End
    End If
    
    ' Clear openDSS before doing anything
    DSSText.Command = "clear"
    ' Compile the script
    DSSText.Command = "compile " + ActiveWorkbook.Path + "\Networks\" + Trim(Network) + "\" + Trim(Network)
    
    ' Initialise Profiles ---------

    Tmonth = Int(ChooseNetwork.MonthVal.Value)
    Tday = Int(ChooseNetwork.Tday)
    
    If Network = "Urban" Then customers = 632
    If Network = "SemiUrban" Then customers = 468
    If Network = "Rural" Then customers = 132
    
    If Start.OverrideDefault = True Then
        DSSText.Command = "Transformer.LV_Transformer.kvs=(11, " & (AdvancedProperties.TransformerVoltage) / 1000 & ")"
    End If
    
    If ChooseNetwork.EVPeneScroll.Value <> 0 Then
        EVPenetration = ChooseNetwork.EVPeneText.Value / 100
        Call Assign_EV_Profiles(customers, EVPenetration)
    End If
    
    If ChooseNetwork.PVPeneScroll.Value <> 0 Then
        PVPenetration = ChooseNetwork.PVPeneText.Value / 100
        location = ChooseNetwork.SelectLocation.ListIndex + 1
        clearness = ChooseNetwork.ClearnessText.Value
        Call Assign_PV_Profiles(customers, PVPenetration, location, Tmonth, clearness)
    End If
    
    If ChooseNetwork.HPPeneScroll.Value <> 0 Then
        HPPenetration = ChooseNetwork.HPPeneText.Value / 100
        location = ChooseNetwork.SelectLocation.ListIndex + 1
        Call Assign_HP_Profiles(customers, HPPenetration, Tmonth, Tday, location)
    End If
    
    If ChooseNetwork.CHPPeneScroll.Value <> 0 Then
        CHPPenetration = ChooseNetwork.CHPPeneText.Value / 100
        location = ChooseNetwork.SelectLocation.ListIndex + 1
        Call Assign_CHP_Profiles(customers, CHPPenetration, Tmonth, Tday, location)
    End If
    
    Call Assign_House_Profiles(customers, Tmonth, Tday)
    
    '------------------------------
    
    DSSText.Command = "Transformer.LV_Transformer.tap=" & (1 + (ChooseNetwork.TransformerTap.Value / 100)) 'Adjust the Off-Load Tap position


End Sub
