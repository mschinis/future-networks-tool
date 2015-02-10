VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomNetworkOptionsForm 
   Caption         =   "Custom network options"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4320
   OleObjectBlob   =   "CustomNetworkOptionsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustomNetworkOptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ContinueButtonPressed_Click()
    Dim networkName As String
    Dim fso As Object
    Dim oFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Options from form
    networkName = CustomNetworkOptionsForm.NetworkNameTextField.Text
    transformerSize = CInt(CustomNetworkOptionsForm.TransformerSizeTextField.Text)
    noOfCustomers = CInt(CustomNetworkOptionsForm.NumberOfCustomersTextField.Text)
    
    strpath = ActiveWorkbook.Path & "\Networks\" & networkName
    CheckDir (strpath)
    
    ' Create the Master file
    Set oFile = fso.CreateTextFile(strpath & "\" & networkName & ".dss")
    
    oFile.WriteLine "Clear"
    oFile.WriteLine "New Circuit." & networkName & "LVNetwork"
    oFile.WriteLine "Edit Vsource.Source BasekV=11 pu=1.00 angle=0 ISC3=3000 ISC1=2500"
    oFile.WriteLine "New transformer.LV_Transformer Buses=(Sourcebus, Main_Busbar)  Conns=(Delta, Wye) kvs=(11, 0.433) kvas=(200, 200) xhl=4.5"
    oFile.WriteLine ""
    oFile.WriteLine "Redirect Linecodes.txt"
    oFile.WriteLine "Redirect " & networkName & "_LinesLaterals1.txt"
    oFile.WriteLine "Redirect " & networkName & "_LinesLaterals2.txt"
    oFile.WriteLine "Redirect " & networkName & "_LinesLaterals3.txt"
    oFile.WriteLine "Redirect " & networkName & "_LinesLaterals4.txt"
    
    oFile.WriteLine "Redirect " & networkName & "_Consumers1.txt"
    oFile.WriteLine "Redirect " & networkName & "_Consumers2.txt"
    oFile.WriteLine "Redirect " & networkName & "_Consumers3.txt"
    oFile.WriteLine "Redirect " & networkName & "_Consumers4.txt"
    
    oFile.WriteLine "Monitors.txt"
    oFile.WriteLine "EnergyMeters.txt"
    oFile.WriteLine ""
    oFile.WriteLine "Set voltagebases=[11 0.4]"
    oFile.WriteLine "CalcVoltageBases"
    
    oFile.Close
    
    ' Create Linecodes File
    Set oFile = fso.CreateTextFile(strpath & "\Linecodes.txt")
    
    oFile.WriteLine "New Linecode.Type-A R1=0.102 X1=0.068 R0=0.625 X0=0.085 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.WriteLine "New Linecode.Type-B R1=0.127 X1=0.073 R0=0.619 X0=0.109 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.WriteLine "New Linecode.Type-C R1=0.166 X1=0.0685 R0=0.625 X0=0.088 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.WriteLine "New Linecode.Type-D R1=0.322 X1=0.069 R0=1.201 X0=0.097 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.WriteLine "New Linecode.Type-E R1=1.2 X1=0.079 R0=1.3 X0=0.079 C0=0.0 C1=0.0 units=km nphases=1"
        
    oFile.WriteLine "New Linecode.Line_185 R1=0.164 X1=0.0685 R0=0.625 X0=0.088 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.WriteLine "New Linecode.Line_95 R1=0.320 X1=0.069 R0=1.201 X0=0.097 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.WriteLine "New Linecode.Line_25 RMATRIX=[1.18] XMATRIX=[0.0515] C=[0.0] units=km   nphases=1"
    
    oFile.Close
    
    ' Create settings file
    Set oFile = fso.CreateTextFile(strpath & "\settings.csv")
    oFile.WriteLine "Customers," & noOfCustomers
    oFile.Close
    
    ' Finish creation of files
    Set fso = Nothing
    Set oFile = Nothing
    
    MsgBox "Files created. To run a simulation on the new network, click 'Start' and then 'Load Generic Network'"
    CustomNetworkOptionsForm.Hide
End Sub

Private Sub FeederLengthSpinButton_SpinDown()
    CustomNetworkOptionsForm.FeederLengthTextField.Text = CStr(CInt(CustomNetworkOptionsForm.FeederLengthTextField) - 10)
End Sub

Private Sub FeederLengthSpinButton_SpinUp()
    CustomNetworkOptionsForm.FeederLengthTextField.Text = CStr(CInt(CustomNetworkOptionsForm.FeederLengthTextField) + 10)
End Sub

Private Sub LateralLengthSpinButton_SpinDown()
    CustomNetworkOptionsForm.LateralLengthTextField.Text = CStr(CInt(CustomNetworkOptionsForm.LateralLengthTextField) - 10)
End Sub

Private Sub LateralLengthSpinButton_SpinUp()
    CustomNetworkOptionsForm.LateralLengthTextField.Text = CStr(CInt(CustomNetworkOptionsForm.LateralLengthTextField) + 10)
End Sub

Private Sub TransformerSizeSpinButton_SpinDown()
    CustomNetworkOptionsForm.TransformerSizeTextField.Text = CStr(CInt(CustomNetworkOptionsForm.TransformerSizeTextField) - 10)
End Sub

Private Sub TransformerSizeSpinButton_SpinUp()
    CustomNetworkOptionsForm.TransformerSizeTextField.Text = CStr(CInt(CustomNetworkOptionsForm.TransformerSizeTextField) + 10)
End Sub
