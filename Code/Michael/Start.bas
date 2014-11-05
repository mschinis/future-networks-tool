Attribute VB_Name = "Start"
Public RunHours As Integer
Public CustomersLimits() As Byte

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


    WelcomeScreen.Show ' Goes into either Preset or Custom Network after this
    If ChooseNetwork.finished <> True Then Exit Sub
    
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
    Dim TransformerArray() As Double
    Dim Feeders() As Double
    Dim Laterals() As Double
    Dim CustomersVoltages() As Double
    ReDim TransformerArray(1 To RunHours, 1 To 4) ' (iteration, 1 = transformerpwoer, 2-4 voltages)
    ReDim Feeders(1 To RunHours, 1 To 4, 1 To 3) ' (iteration, feeder, currentstarts)
    ReDim Laterals(1 To RunHours, 1 To 4, 1 To 4, 1 To 9) ' (iteration, feeder, lateral, 1-9 currents / voltagesstart / voltagesend)
    ReDim CustomersVoltages(1 To 4, 1 To (PresetNetwork.customers / 4), 1 To RunHours)
    ReDim CustomersLimits(1 To 4, 1 To (PresetNetwork.customers / 4), 1 To RunHours)


    
    For i = 1 To RunHours
    
        DSSobj.ActiveCircuit.Solution.Solve
        Call CheckValuesPreset(PresetNetwork.customers, i, TransformerArray, Feeders, Laterals, CustomersVoltages, CustomersLimits)
        
    Next
    
    
    Call Check_Compliance
    
    DSSText.Command = "Export monitors Transformer"
    DSSText.Command = "Export meters"
' ----- end coding here -----
   
   
   Call Monitors
   
    
    MsgBox ("Total time " + Trim(Str(Timer - stime)))
End Sub
