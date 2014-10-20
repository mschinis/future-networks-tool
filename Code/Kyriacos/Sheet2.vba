VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' gamw ton kaiju

Private Sub Worksheet_Activate()
    ' Populate the combobox with the available networks from the Settings worksheet
    With Sheet3
     Sheet2.ComboBox1.List = .Range("B4:B" & .Range("B" & .Rows.Count).End(xlUp).Row).Value
    End With
    
    
End Sub
' ----------------- Start OpenDSS Button -------------------------
Private Sub Start_OpenDSS_Click()
    
    ' Create a new instance of the DSS
    Set DSSobj = New OpenDSSengine.DSS
           
    ' Start the DSS
    If Not DSSobj.Start(0) Then
        MsgBox "DSS Failed to Start"
    Else
        MsgBox "DSS Started successfully"
        ' Assign a variable to the Text interface for easier access
        Set DSSText = DSSobj.Text
    End If
End Sub
' ----------------- End OpenDSS Button -------------------------

' ----------------- Start Solve Button -------------------------
Private Sub Solve_Button_Click()
    Dim stime As Single
    Dim EVPenetration, PVPenetration As Double
    Dim location, Tmonth, Tday, clearness As Integer
    Dim network As String
    
    WelcomeScreen.Show
    network = ChooseNetwork.SelectNetwork.Value ' Select Network from Dropdown Menu

    
    ' Clear openDSS before doing anything
    DSSText.Command = "clear"
    ' Compile the script
    DSSText.Command = "compile " + ActiveWorkbook.Path + "\Networks\" + network + "\" + Trim(network)
    
    ' Initialise Profiles ---------

    Tmonth = ChooseNetwork.MonthVal.Value
    Tday = ChooseNetwork.Tday
    Call Assign_House_Profiles(632, Tmonth, Tday)
    
    
    If ChooseNetwork.EVEnable.Value = True Then
        EVPenetration = ChooseNetwork.EVPeneText.Value / 100
        Call Assign_EV_Profiles(632, EVPenetration)
    End If
    
    If ChooseNetwork.PVEnable.Value = True Then
        PVPenetration = ChooseNetwork.PVPeneText.Value / 100
        location = ChooseNetwork.SelectNetwork.ListIndex + 1
        clearness = ChooseNetwork.ClearnessText.Value
        'Call Assign_PV_Profiles(632, PVPenetration, location, Tmonth, clearness)
    End If
    

    
    
    
    
    
    '------------------------------
    DSSText.Command = "Set Datapath =" & ActiveWorkbook.Path & "\output"
    DSSText.Command = "new monitor.Transformer element=transformer.LV_Transformer terminal=1 mode=1 ppolar=yes"
    
    
    
   
    ' The Compile command sets the current directory the that of the file
    ' Thats where all the result files will end up.
    
    ' Assign a variable to the Circuit interface for easier access
    Set DSSCircuit = DSSobj.ActiveCircuit
    Set DSSSolution = DSSCircuit.Solution
    Set DSSControlQueue = DSSCircuit.CtrlQueue
    
    runHours = Worksheets("Settings").Range("D5").Value
            
    stime = Timer
    
    DSSText.Command = "Set ControlMode=time"
    DSSText.Command = "Reset" 'resetting all energy meters and monitors
    DSSobj.AllowForms = False 'no "solution progress" window
    DSSText.Command = "Set Mode=daily stepsize=1m number=1"
    
' ----- start coding here -----
    Dim i As Integer
    

    For i = 1 To runHours
    
        DSSobj.ActiveCircuit.Solution.Solve
        
    Next

    DSSText.Command = "Export monitors Transformer"
' ----- end coding here -----

    
    MsgBox ("Total time " + Trim(Str(Timer - stime)))
End Sub

' ----------------- End Solve Button -------------------------

Private Sub CommandButton7_Click()
    
    runHours = Range("E8").Value
    Monitors
    
    MsgBox ("Done")
    
End Sub

Private Sub ComboBox29_Change()
    MsgBox ("complete")
    
End Sub
