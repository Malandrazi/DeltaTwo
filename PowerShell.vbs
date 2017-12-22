
Const SVSFlagsAsync = 1

Dim Speech
Dim FSO
Dim max,min
max=10
min=0
Randomize
CreateObjects
Main
DestroyObjects
Quit

Sub Main
        Dim sText1
		Dim sText2
		Dim sText3
		Dim sText4
		Dim sText5
		Dim sText6
		Dim sText7
		Dim sText8
		Dim sText9
		Dim sText0
			
		sText1 ="Sir,I have detected power shell opening"
		sText2 ="Did you start Power shell Sir?"
		sText3 ="Power shell is running Sir"
		sText4 ="Starting power shell Sir"
		sText5 ="Power shell console detected, sir"
		sText6 ="Detecting a power shell console, sir"
		sText7 ="Sir, should I be starting power shell?"
		sText8 ="Power shell has started"
		sText9 ="Sir, are you running power shell?"
		sText0 ="Power shell is started"
		
		
       sText = (Int((max-min)*Rnd+min))
        sText = Trim(sText)
		 If sText = "1" Then
                SpeakText sText1
				else
        If sText = "2" Then
                SpeakText sText2
				else
		 If sText = "3" Then
                SpeakText sText3
				else
		If sText = "4" Then
                SpeakText sText4
				else
		 If sText = "5" Then
                SpeakText sText5
				else
		If sText = "6" Then
                SpeakText sText6
				else
        If sText = "7" Then
                SpeakText sText7
				else
		 If sText = "8" Then
                SpeakText sText8
				else
		If sText = "9" Then
                SpeakText sText9
				else
		 If sText = "0" Then
                SpeakText sText0
		
		End If
		End If		
        End If
		End If
		End If
		End If
		End If		
        End If
		End If
		End If
End Sub

Sub SpeakText(sText)
        On Error Resume Next
        Speech.Speak sText, SVSFlagsAsync 
        Do
                Sleep 100
        Loop Until Speech.WaitUntilDone(10)
End Sub

Sub StopSpeaking()
        On Error Resume Next
        Speech.Speak vbNullString
        Set Speech = Nothing
End Sub

Sub CreateObjects
        Set Speech = CreateObject("SAPI.SpVoice")
        Set FSO = CreateObject("Scripting.FileSystemObject")
End Sub

Sub DestroyObjects
        Set Speech = Nothing
        Set FSO = Nothing
End Sub

Sub Sleep(nTimeout)
        WScript.Sleep nTimeout
End Sub

Sub Quit
        WScript.Quit
End Sub