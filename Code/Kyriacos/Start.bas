Attribute VB_Name = "Start"
Public RunHours As Integer
Public CustomersLimits() As Byte
Public CustomerVoltageLimit() As Byte
Public OverrideDefault As Boolean
Public CurrentFlags() As Byte
Public NotCompliant() As Double

Public TransformerArray() As Double
Public Feeders() As Double
Public Laterals() As Double


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

OverrideDefault = False


'''''''''''''''''''''''''''''''''''''
'    Dim StatusOld As Boolean, CalcOld As XlCalculation
'
'    ' Capture Initial Settings
'    StatusOld = Application.DisplayStatusBar
'
'    '      Doing these will speed up your code
'    CalcOld = Application.Calculation
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

    ChooseNetwork.finished = False
    WelcomeScreen.Show ' Goes into either Preset or Custom Network after this
    If ChooseNetwork.finished <> True Then GoTo ENDLINE
    
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
    ReDim Feeders(1 To RunHours, 1 To 4, 1 To 3) ' (iteration, feeder, currentstarts)
    ReDim Laterals(1 To RunHours, 1 To 4, 1 To 4, 1 To 9) ' (iteration, feeder, lateral, 1-9 currents / voltagesstart / voltagesend)
    ReDim CustomersVoltages(1 To 4, 1 To (PresetNetwork.customers / 4), 1 To RunHours)
    ReDim CustomersLimits(1 To 4, 1 To (PresetNetwork.customers / 4), 1 To RunHours)
    ReDim CustomerVoltageLimit(1 To PresetNetwork.customers)
    ReDim CurrentFlags(1 To 4, 1 To 5)
    ReDim NotCompliant(1 To PresetNetwork.customers)
    
    progresscounter = 0
    Application.StatusBar = "Simulation running - 10%"
    
    For i = 1 To 1860
           
        DSSobj.ActiveCircuit.Solution.Solve
        If i > 420 Then
            If i Mod 180 = 0 Then
                progresscounter = progresscounter + 1
                Application.StatusBar = "Simulation running - " & (progresscounter * 10 + 10) & "%"
            End If
            Call CheckValuesPreset(PresetNetwork.customers, i - 420, TransformerArray, Feeders, Laterals, CustomersVoltages, CustomersLimits, CurrentFlags)
            If ChooseNetwork.EVPeneScroll.Value <> 0 Then
                Call EVManagement(i - 420)
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
