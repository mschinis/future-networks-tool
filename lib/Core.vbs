dim fso
dim curDir
dim WinScriptHost
set fso = CreateObject("Scripting.FileSystemObject")
curDir = fso.GetAbsolutePathName(".")
set fso = nothing

Set xlObj = CreateObject("Excel.application")
xlObj.Workbooks.Open curDir & "\Start Here.xlsm"
xlObj.Run "Start.Start"