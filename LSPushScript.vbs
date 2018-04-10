Set WshShell = CreateObject("Wscript.Shell")
WshShell.run "%logonserver%\netlogon\LsPush.exe 10.128.4.235",0