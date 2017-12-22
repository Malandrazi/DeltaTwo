' Test program for the IncludeFile and ReadConfigFile functions.
' Author: Christian d'Heureuse (www.source-code.biz)
' License: GNU/LGPL (http://www.gnu.org/licenses/lgpl.html)

Option Explicit

Dim fso: set fso = CreateObject("Scripting.FileSystemObject")

' Includes a file in the global namespace of the current script.
' The file can contain any VBScript source code.
' The path of the file name must be specified relative to the
' directory of the main script file.
Private Sub IncludeFile (ByVal RelativeFileName)
   Dim ScriptDir: ScriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
   Dim FileName: FileName = fso.BuildPath(ScriptDir,RelativeFileName)
   IncludeFileAbs FileName
   End Sub

' Includes a file in the global namespace of the current script.
' The file can contain any VBScript source code.
' The path of the file name must be specified absolute (or
' relative to the current directory).
Private Sub IncludeFileAbs (ByVal FileName)
   Const ForReading = 1
   Dim f: set f = fso.OpenTextFile(FileName,ForReading)
   Dim s: s = f.ReadAll()
   ExecuteGlobal s
   End Sub

' Includes the configuration file.
' The configiguration file has the name of the main script
' with the extension ".config".
Private Sub ReadConfigFile
   Dim ConfigFileName: ConfigFileName = fso.GetBaseName(WScript.ScriptName) & ".config"
   IncludeFile ConfigFileName
   End Sub

ReadConfigFile

WScript.Echo "I added a new variable=" & NewVar
WScript.Echo "I can change this=" & ConfigParm1
WScript.Echo "this is just label text and can have spaces but " & canthavespaces
WScript.Echo "I can make a script that teaches how to script" & scriptteacher

Sub1
'Sub1 code is stored in the variables file

'run this script from the original by adding the appropriate code to the Original config file
