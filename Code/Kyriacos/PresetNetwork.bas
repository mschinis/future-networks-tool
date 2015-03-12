Attribute VB_Name = "PresetNetwork"
Public Network As String
Public customers As Integer

Public PVEnabled As Boolean
Public EVEnabled As Boolean
Public HPEnabled As Boolean
Public CHPEnabled As Boolean

Public Parser As ParserXControl.ParserX

Public Sub Preset_Network()

Dim stime As Single
    Dim EVPenetration, PVPenetration, HPPenetration, CHPPenetration As Double
    Dim location, Tmonth, Tday, clearness As Integer
    Network = ChooseNetwork.SelectNetwork.Value ' Select Network from Dropdown Menu
        
    Dim File_Location As String
    
    Dim FileNum As Long
    Dim parserExtraStr As String
    Dim s As String
    
    ReDim Assign_Profiles.CHPStopPoint(1 To Assign_Profiles.NoLaterals, 1 To Assign_Profiles.NoFeeders)
    ReDim Assign_Profiles.HPStopPoint(1 To Assign_Profiles.NoLaterals, 1 To Assign_Profiles.NoFeeders)
    
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
    
    ' Setup parser
    Set Parser = Nothing ' destroy old object should it already exist
    Set Parser = New ParserXControl.ParserX
    Parser.AutoIncrement = True
    FileNum = FreeFile
    
    ' Find the number of customers from the settings.csv file of the network
    Open ActiveWorkbook.Path & "\Networks\" & Network & "\settings.csv" For Input As #FileNum
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        customers = Parser.IntValue
    Close
    
    If Start.OverrideDefault = True Then
        DSSText.Command = "Transformer.LV_Transformer.kvs=(11, " & (AdvancedProperties.TransformerVoltage) / 1000 & ")"
    End If
    
    If ChooseNetwork.EVEnable.Value = True Then
        EVPenetration = ChooseNetwork.EVPeneText.Value / 100
        Call Assign_EV_Profiles(customers, EVPenetration)
    End If
    
    If ChooseNetwork.PVEnable.Value Then
        PVPenetration = ChooseNetwork.PVPeneText.Value / 100
        location = ChooseNetwork.SelectLocation.ListIndex + 1
        clearness = ChooseNetwork.ClearnessText.Value
        Call Assign_PV_Profiles(customers, PVPenetration, location, Tmonth, clearness)
    End If
    
    If ChooseNetwork.HPEnable.Value = True Then
        HPPenetration = ChooseNetwork.HPPeneText.Value / 100
        location = ChooseNetwork.SelectLocation.ListIndex + 1
        Call Assign_HP_Profiles(customers, HPPenetration, Tmonth, Tday, location)
    End If
    
    If ChooseNetwork.CHPEnable.Value = True Then
        CHPPenetration = ChooseNetwork.CHPPeneText.Value / 100
        location = ChooseNetwork.SelectLocation.ListIndex + 1
        Call Assign_CHP_Profiles(customers, CHPPenetration, Tmonth, Tday, location)
    End If
    
    Call Assign_House_Profiles(customers, Tmonth, Tday)
    
    '------------------------------
    
    DSSText.Command = "Transformer.LV_Transformer.tap=" & (1 + (ChooseNetwork.TransformerTap.Value / 100)) 'Adjust the Off-Load Tap position


End Sub
