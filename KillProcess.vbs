' Script para eliminar procesos en Windows, requiere una variable de entrada desde Automation AnyWhere. Variable de entrada deberá incluir el nombre y extension: notepad.exe

Dim strProcessKill, objWMIService, colProcess, strComputer, strList, p
strProcessKill = WScript.Arguments.Item(0)

WScript.Sleep 2000

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2") 
Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name like '" & strProcessKill & "'")
For Each p in colProcess
  p.Terminate             
Next