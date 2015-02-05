Attribute VB_Name = "VersionControl"
Sub SaveCodeModules()

'This code Exports all VBA modules
Dim i%, sName$

With ThisWorkbook.VBProject
    For i% = 1 To .VBComponents.Count
        If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
            sName$ = .VBComponents(i%).CodeModule.name
            .VBComponents(i%).Export "C:\Users\michael\Documents\future-networks-tool\Code\Michael\" & sName$ & ".bas"
        End If
    Next i
End With

End Sub
