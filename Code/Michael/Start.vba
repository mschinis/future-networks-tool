Attribute VB_Name = "Start"
Public RunHours As Integer

Public Sub Start()
' Create a new instance of the DSS
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
    

    For i = 1 To RunHours
    
        DSSobj.ActiveCircuit.Solution.Solve
        
    Next

    DSSText.Command = "Export monitors Transformer"
' ----- end coding here -----

    Call Monitors
    MsgBox ("Total time " + Trim(Str(Timer - stime)))
    
End Sub

