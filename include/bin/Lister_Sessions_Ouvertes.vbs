strComputer = WScript.Arguments(0)

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 
For Each objComputer in colComputer
    Wscript.Echo "Utilisateur connecté localement : " & objComputer.UserName
Next
