Attribute VB_Name = "SharedClass"
' Dat jizz singleton tho
Private settingsSharedClass As SimulationSettings
Public Function Settings() As SimulationSettings
    If settingsSharedClass Is Nothing Then
        Set settingsSharedClass = New SimulationSettings
    End If
    Set Settings = settingsSharedClass
End Function
Public Function resetSettings()
    Set settingsSharedClass = Nothing
End Function
