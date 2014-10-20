Attribute VB_Name = "PresetNetwork"
Public Network As String

Public Sub Preset_Network()

Dim stime As Single
    Dim EVPenetration, PVPenetration As Double
    Dim location, Tmonth, Tday, clearness, customers As Integer
    Network = ChooseNetwork.SelectNetwork.Value ' Select Network from Dropdown Menu
    Dim File_Location As String
    
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
    
    
    Call Assign_House_Profiles(customers, Tmonth, Tday)
    
    
    If ChooseNetwork.EVEnable.Value = True Then
        EVPenetration = ChooseNetwork.EVPeneText.Value / 100
        Call Assign_EV_Profiles(customers, EVPenetration)
    End If
    
    If ChooseNetwork.PVEnable.Value = True Then
        PVPenetration = ChooseNetwork.PVPeneText.Value / 100
        location = ChooseNetwork.SelectLocation.ListIndex + 1
        clearness = ChooseNetwork.ClearnessText.Value
        'Call Assign_PV_Profiles(customers, PVPenetration, location, Tmonth, clearness)
    End If
    

    '------------------------------
    


End Sub
