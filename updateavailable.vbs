
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
			
		sText1 ="Sir, an update has been detected"
		sText2 ="Sir Update Detected"
		sText3 ="Sir,I've detected an update which is available"
		sText4 ="Sir,I have detected an available update"
		sText5 ="I'm detecting an available update"
		sText6 ="Check for available updates, Sir"
		sText7 ="Check Windows Update Sir, one is available"
		sText8 ="an available update can be downloaded Sir"
		sText9 ="Sir,I have found an update"
		sText0 ="Sir,an update has been found"
		
		
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