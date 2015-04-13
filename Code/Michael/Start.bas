Attribute VB_Name = "Start"
Public RunHours As Integer
Public CustomersLimits() As Byte
Public CustomerVoltageLimit() As Byte
Public OverrideDefault As Boolean
Public CurrentFlags() As Byte
Public NotCompliant() As Double

Public TransformerArray() As Double
Public feeders() As Double
Public laterals() As Double

Public finished As Boolean


Public Parser As ParserXControl.ParserX

Public Sub getSettings()
    '-------------------------- Load Settings File in Singleton ---------------------------------------------------
    Dim FileNum As Long
    Dim parserExtraStr As String
    Dim s As String
    SharedClass.resetSettings
    Dim localObjSettings As SimulationSettings
    Set localObjSettings = SharedClass.Settings
    
    Set Parser = Nothing ' destroy old object should it already exist
    Set Parser = New ParserXControl.ParserX
    Parser.AutoIncrement = True
    FileNum = FreeFile
    
    localObjSettings.network = ChooseNetwork.SelectNetwork.Value
    
    ' Load all settings from the settings.csv file of the network
    Open ActiveWorkbook.Path & "\Networks\" & localObjSettings.network & "\settings.csv" For Input As #FileNum
        ' Customers
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        localObjSettings.customers = Parser.IntValue
        ' Feeders
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        localObjSettings.feeders = Parser.IntValue
        ' Laterals
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        localObjSettings.laterals = Parser.IntValue
        ' Transformer Size
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        localObjSettings.transformerSize = Parser.IntValue
        ' Feeder Winter Current Limit
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        localObjSettings.feederWinterCurrentLimit = Parser.IntValue
        ' Feeder Summer Current Limit
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        localObjSettings.feederSummerCurrentLimit = Parser.IntValue
        ' Lateral Winter Current Limit
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        localObjSettings.lateralWinterCurrentLimit = Parser.IntValue
        ' Lateral Summer Current Limit
        Line Input #FileNum, s
        Parser.CmdString = s
        parserExtraStr = Parser.StrValue
        localObjSettings.lateralSummerCurrentLimit = Parser.IntValue
    Close
End Sub

Public Sub Start()

'initialise values with imposibru values

CheckValues.MaxCurrentUseLateral = 0
CheckValues.MinCurrentUseLateral = 10
CheckValues.MaxTransformerUse = 0
CheckValues.MinTransformerUse = 10
CheckValues.MaxVoltage = 0
CheckValues.MinVoltage = 2
CheckValues.MaxCurrentUseFeeder = 0
CheckValues.MinCurrentUseFeeder = 10
Assign_Profiles.HPEnabled = False
Assign_Profiles.CHPEnabled = False

OverrideDefault = False

'''''''''''''''''''''''''''''''''''''
'    Dim StatusOld As Boolean, CalcOld As XlCalculation
'
'    ' Capture Initial Settings
'    StatusOld = Application.DisplayStatusBar
'
'    '      Doing these will speed up your code
    CalcOld = Application.Calculation
'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'
'    On Error GoTo EH
'
'    Application.StatusBar = "Simulation running - 0%"
''
'''''''''''''''''''''''''



'Sheets("Network").Activate
'For em = 1 To 4
'    For emm = 1 To 5
'            Sheets("Network").Shapes("Feeder" & em & "Lateral" & emm - 1).Visible = False
'    Next
'Next
'Dim Shp As Shape
'For Each Shp In ActiveSheet.Shapes
'    If Shp.Type = 1 Then
'        Shp.Delete
'    End If
'Next Shp
'Sheets("Main").Activate

' Create a new instance of the DSS
    Reset
    Set DSSobj = New OpenDSSengine.DSS
           
    ' Start the DSS
    If Not DSSobj.Start(0) Then
        MsgBox "DSS Failed to Start"
        Exit Sub
    Else

        Set DSSText = DSSobj.Text
    End If
    
    finished = False
    WelcomeScreen.Show ' Goes into either Preset or Custom Network after this
    If finished <> True Then GoTo ENDLINE
    
    DSSText.Command = "Set Datapath =" & ActiveWorkbook.Path & "\output"
    DSSText.Command = "new monitor.Transformer element=transformer.LV_Transformer terminal=1 mode=1 ppolar=yes"
    
    
    ' The Compile command sets the current directory the that of the file
    ' Thats where all the result files will end up.
    
    ' Assign a variable to the Circuit interface for easier access
    Set DSSCircuit = DSSobj.ActiveCircuit
    Set DSSSolution = DSSCircuit.Solution
    Set DSSControlQueue = DSSCircuit.CtrlQueue
    
    RunHours = 1440
            
    stime = Timer
    
    DSSText.Command = "Set ControlMode=time"
    DSSText.Command = "Reset" 'resetting all energy meters and monitors
    DSSobj.AllowForms = False 'no "solution progress" window
    DSSText.Command = "Set Mode=daily stepsize=1m number=1"
    
' ----- start coding here -----

    Dim i As Integer
    Dim CustomersVoltages() As Double
    ReDim TransformerArray(1 To RunHours, 1 To 4) ' (iteration, 1 = transformerpwoer, 2-4 voltages)
    ReDim feeders(1 To RunHours, 1 To Assign_Profiles.NoFeeders, 1 To 3) ' (iteration, feeder, currentstarts)
    ReDim laterals(1 To RunHours, 1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals, 1 To 9) ' (iteration, feeder, lateral, 1-9 currents / voltagesstart / voltagesend)
    ReDim CustomersVoltages(1 To Assign_Profiles.NoFeeders, 1 To (PresetNetwork.customers / Assign_Profiles.NoFeeders), 1 To RunHours)
    ReDim CustomersLimits(1 To Assign_Profiles.NoFeeders, 1 To (PresetNetwork.customers / Assign_Profiles.NoFeeders), 1 To RunHours)
    ReDim CustomerVoltageLimit(1 To PresetNetwork.customers)
    ReDim CurrentFlags(1 To Assign_Profiles.NoFeeders, 1 To Assign_Profiles.NoLaterals + 1)
    ReDim NotCompliant(1 To PresetNetwork.customers)
    
    progresscounter = 0
    Application.StatusBar = "Simulation running - 10%"
    
    For i = 1 To RunHours + 420
           
        DSSobj.ActiveCircuit.Solution.Solve
        If i > 420 Then
            If i Mod 180 = 0 Then
                progresscounter = progresscounter + 1
                Application.StatusBar = "Simulation running - " & (progresscounter * 10 + 10) & "%"
            End If
            Call CheckValuesPreset(PresetNetwork.customers, i - 420, TransformerArray, feeders, laterals, CustomersVoltages, CustomersLimits, CurrentFlags)
            If ChooseNetwork.EVEnable.Value = True Then
                Call EVManagement(i - 420)
            End If
            
            If ChooseNetwork.PVANM.Value = True Then
                Call PVManagement(i - 420)
            End If
            
            If ChooseNetwork.HPANM.Value = True Then
                Call HPManagement(i - 420)
            End If
            
        End If
    Next



    Call Check_Compliance
    Call Customer_Voltage_Percentage
    
    DSSText.Command = "Export monitors Transformer"
    DSSText.Command = "Export meters"
' ----- end coding here -----
   
   
   Call Monitors
   
    Application.StatusBar = "Simulation running - 100%"
    'CostCalculations.CalculateCosts
    MsgBox ("Total time " + Trim(Str(Timer - stime)))
    
ENDLINE:
    ActiveWorkbook.RefreshAll
    Application.StatusBar = False
    Application.Calculation = CalcOld
    Application.DisplayStatusBar = StatusOld
    Application.ScreenUpdating = True
    Application.EnableEvents = True

EH:

    'Error handler


End Sub
