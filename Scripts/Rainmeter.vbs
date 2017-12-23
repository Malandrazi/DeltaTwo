Sub Run(ByVal sFile)
Dim shell

    Set shell = CreateObject("WScript.Shell")
    shell.Run Chr(34) & sFile & Chr(34), 1, false
    Set shell = Nothing
End Sub
Run "C:\Program Files (x86)\SpeedFan\speedfan.exe"
Run "C:\Program Files\Rainmeter\Rainmeter.exe"
