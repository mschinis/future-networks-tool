VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomNetworkOptionsForm 
   Caption         =   "Custom network options"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4305
   OleObjectBlob   =   "CustomNetworkOptionsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "CustomNetworkOptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Things to refactor
' CalculateLateral.LateralNo
' CalculateLateral.FeederLength
' CalculateLateral.lateralLength
'
' CheckValues.CheckValuesPreset
' Sheet2.CommandButton1_Click - Draw Network

Private Sub ContinueButtonPressed_Click()
    Dim fso As Object
    Dim oFile As Object
    Dim i, j As Integer
    Dim k, l As Long
    Dim customersPerFeederPerLateral() As Variant
    Dim customersPerLaterals() As Variant
    Dim lateralStarts() As Long
    Dim lateralEnds() As Long
    
    Dim customerEquidistance As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Dim all userForm values
    Dim networkName As String
    Dim noOfFeeders As Integer
    Dim noOfLaterals As Integer
    Dim noOfCustomers As Long
    Dim transformerSize As Integer
    Dim feederLength As Integer
    Dim lateralLength As Integer
    
    ' Options from Userform
    networkName = CustomNetworkOptionsForm.NetworkNameTextField.Text
    noOfFeeders = Int(CustomNetworkOptionsForm.NumberOfFeedersTextField.Text)
    noOfLaterals = Int(CustomNetworkOptionsForm.NumberOfLateralsTextField.Text)
    noOfCustomers = CLng(CustomNetworkOptionsForm.NumberOfCustomersTextField.Text)
    transformerSize = Int(CustomNetworkOptionsForm.TransformerSizeTextField.Text)
    ' Feeder length not implemented yet
    lateralLength = Int(CustomNetworkOptionsForm.LateralLengthTextField.Text)
    
    distanceFromFirstLateral = 50
    equidistanceFromLaterals = 100
    
    ' Copy the number of customers in a temporary storage
    noOfCustomersTemp = noOfCustomers
    ' Determine the number of customers per Feeder and Lateral
    noOfCustomersPerFeeder = Int(noOfCustomers / noOfFeeders)
    noofcustomersperlateral = Int(noOfCustomersPerFeeder / noOfLaterals)
    
    
    ' Validation of input
    singleSpace = " "
    For i = 1 To Len(networkName)
        If Mid(networkName, i, 1) = singleSpace Then
            MsgBox "Spaces in the network name not allowed."
            Exit Sub
        End If
    Next i
    If noOfFeeders = 0 Then
        MsgBox "There must be allocated at least one feeder on the network"
        Exit Sub
    End If
    If noOfLaterals = 0 Then
        MsgBox "There must be allocated at least one lateral on the network"
        Exit Sub
    End If
    If noOfCustomers = 0 Then
        MsgBox "There must be allocated at least 1 customer on the network"
        Exit Sub
    End If
    If transformerSize = 0 Then
        MsgBox "Transformer size cannot be zero"
        Exit Sub
    End If
    If lateralLength < 20 Then
        MsgBox "Lateral length must be at least 20m"
        Exit Sub
    End If
    If noofcustomersperlateral > lateralLength Then
        MsgBox "Maximum number of customers per lateral (" & noofcustomersperlateral & ") cannot exceed the length of the feeder specified."
        Exit Sub
    End If


    ReDim customersPerFeederPerLateral(1 To noOfFeeders)
    ReDim lateralStarts(1 To noOfLaterals)
    ReDim lateralEnds(1 To noOfLaterals)
    
    ' Alocation of customers on each lateral of each feeder
    For i = 1 To noOfFeeders
        ReDim customersPerLaterals(1 To noOfLaterals)
        For j = 1 To noOfLaterals
            customersPerLaterals(j) = noofcustomersperlateral
            noOfCustomersTemp = noOfCustomersTemp - noofcustomersperlateral
        Next j
        customersPerFeederPerLateral(i) = customersPerLaterals
    Next i
    ' If any customers are not allocated, allocate them on the last lateral of each feeder
    Do While noOfCustomersTemp > 0
        i = 1
        Do While noOfCustomersTemp > 0 And i <= noOfFeeders
            customersPerFeederPerLateral(i)(noOfLaterals) = customersPerFeederPerLateral(i)(noOfLaterals) + 1
            noOfCustomersTemp = noOfCustomersTemp - 1
            i = i + 1
        Loop
    Loop
    
    strpath = ActiveWorkbook.Path & "\Networks\" & networkName
    CheckDir (strpath)
    
    ' Create the Master file
    Set oFile = fso.CreateTextFile(strpath & "\" & networkName & ".dss")
    
    oFile.writeLine "Clear"
    oFile.writeLine "New Circuit." & networkName & "LVNetwork"
    oFile.writeLine "Edit Vsource.Source BasekV=11 pu=1.00 angle=0 ISC3=3000 ISC1=2500"
    oFile.writeLine "New transformer.LV_Transformer Buses=(Sourcebus, Main_Busbar)  Conns=(Delta, Wye) kvs=(11, 0.433) kvas=(200, 200) xhl=4.5"
    oFile.writeLine ""
    oFile.writeLine "Redirect Linecodes.txt"
    
    ' Context of i: i is the number of feeders on the network
    For i = 1 To noOfFeeders
        oFile.writeLine "Redirect " & networkName & "_LinesLaterals" & i & ".txt"
    Next i
    For i = 1 To noOfFeeders
        oFile.writeLine "Redirect " & networkName & "_Consumers" & i & ".txt"
    Next i
    
'    oFile.writeLine "Monitors.txt"
'    oFile.writeLine "EnergyMeters.txt"
    oFile.writeLine ""
    oFile.writeLine "Set voltagebases=[11 0.4]"
    oFile.writeLine "CalcVoltageBases"
    
    oFile.Close
    
    ' Create Linecodes File
    Set oFile = fso.CreateTextFile(strpath & "\Linecodes.txt")
    
    oFile.writeLine "New Linecode.Type-A R1=0.102 X1=0.068 R0=0.625 X0=0.085 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.writeLine "New Linecode.Type-B R1=0.127 X1=0.073 R0=0.619 X0=0.109 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.writeLine "New Linecode.Type-C R1=0.166 X1=0.0685 R0=0.625 X0=0.088 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.writeLine "New Linecode.Type-D R1=0.322 X1=0.069 R0=1.201 X0=0.097 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.writeLine "New Linecode.Type-E R1=1.2 X1=0.079 R0=1.3 X0=0.079 C0=0.0 C1=0.0 units=km nphases=1"
    
    oFile.writeLine "New Linecode.Line_185 R1=0.164 X1=0.0685 R0=0.625 X0=0.088 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.writeLine "New Linecode.Line_95 R1=0.320 X1=0.069 R0=1.201 X0=0.097 C0=0.0 C1=0.0 units=km nphases=3"
    oFile.writeLine "New Linecode.Line_25 RMATRIX=[1.18] XMATRIX=[0.0515] C=[0.0] units=km   nphases=1"
    
    oFile.Close
    
    
        
    For i = 1 To noOfFeeders
        ' Context of i,j,k:
        ' i is the current feeder
        ' j is the current lateral
        ' k is the pointer at the current location
        ' l is the pointer at the current last known feeder end
        l = distanceFromFirstLateral
        
        
        Set oFile = fso.CreateTextFile(strpath & "\" & networkName & "_LinesLaterals" & i & ".txt")
        oFile.writeLine "New Line.Feeder" & i & ".1      Bus1=Main_Busbar Bus2=" & i & "_1       Length=1    units=m Linecode=Line_185"
        ' Allocation of first part of the feeder
        For j = 1 To l
            oFile.writeLine "New Line.Feeder" & i & "." & j + 1 & " Bus1=" & i & "_" & j & "   Bus2=" & i & "_" & j + 1 & "   Length=1    units=m Linecode=Line_185"
        Next j
        k = j
        l = k
        ' Creation of the laterals and rest of the feeder
        For j = 1 To noOfLaterals
            ' Create the lateral
            lateralStarts(j) = k
            oFile.writeLine "New Line.Lateral" & i & "_start_" & j & "   Bus1=" & i & "_" & l - 1 & "   Bus2=" & i & "_" & k + 1 & "   Length=1    units=m Linecode=Line_95"
            k = k + 2
            For k = k To k + lateralLength - 3
                oFile.writeLine "New Line.Lateral" & i & "." & k & "    Bus1=" & i & "_" & k - 1 & "   Bus2=" & i & "_" & k & "   Length=1    units=m Linecode=Line_95"
            Next k
            ' End of lateral
            lateralEnds(j) = k
            oFile.writeLine "New Line.Lateral" & i & "_end_" & j & " Bus1=" & i & "_" & k - 1 & " Bus2=" & i & "_" & k & " units=m Linecode=Line_95 Length=1"
            ' If it's not the last lateral, create the following part of the feeder
            If j <> noOfLaterals Then
                k = k + 1
                oFile.writeLine "New Line.Feeder" & i & "." & k & " Bus1=" & i & "_" & l & "  Bus2=" & i & "_" & k & " units=m  Linecode=Line_185 Length=1 "
                For k = k + 1 To k + equidistanceFromLaterals
                    oFile.writeLine "New Line.Feeder" & i & "." & k & " Bus1=" & i & "_" & k - 1 & "  Bus2=" & i & "_" & k & " units=m  Linecode=Line_185 Length=1 "
                Next k
                k = k - 1
                l = k
            End If
        Next j
        oFile.Close
        
        ' For each of the feeder, allocate the customers on the lateral
        lateralPosition = lateralStarts(1)
        Set oFile = fso.CreateTextFile(strpath & "\" & networkName & "_Consumers" & i & ".txt")
        
        ' Foreach of the laterals, except the last one, allocate their service line
        counter = 1
        For j = 1 To noOfLaterals
            If customersPerFeederPerLateral(i)(j) <> 0 Then
                customerEquidistance = Int((lateralEnds(j) - lateralStarts(j)) / customersPerFeederPerLateral(i)(j))
                For k = 1 To customersPerFeederPerLateral(i)(j)
                    z = k Mod 3
                    If z = 0 Then z = 3
                    
                    l = (customerEquidistance * k) + lateralStarts(j)
                    If l > lateralEnds(j) Then l = lateralEnds(j)
                    oFile.writeLine "New Line.Consumer" & i & "_" & counter & " Bus1=" & i & "_" & l & "." & z & " Bus2=Consumer" & i & "_" & counter & ".1 Length=0.04 units=km Linecode=Line_25"
                    counter = counter + 1
                Next k
            End If
        Next j
        oFile.Close
    Next i
    
    ' Create settings file
    Set oFile = fso.CreateTextFile(strpath & "\settings.csv")
        oFile.writeLine "Customers," & noOfCustomers
        oFile.writeLine "Feeders," & noOfFeeders
        oFile.writeLine "Laterals," & noOfLaterals
        oFile.writeLine "TransformerSize," & transformerSize
        oFile.writeLine "FeederWinterCurrentLimit,"
        oFile.writeLine "FeederSummerCurrentLimit,"
        oFile.writeLine "LateralWinterCurrentLimit,"
        oFile.writeLine "LateralSummerCurrentLimit,"
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
