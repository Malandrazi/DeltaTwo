
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
			
		sText1 ="Sir,I have deleted an exception to the Firewall"
		sText2 ="Did you change the Firewall Sir?"
		sText3 =" deleted an exception from the firewall,Sir"
		sText4 ="deleted an exception from the Firewall Sir"
		sText5 ="a Firewall exception delete detected, sir"
		sText6 ="Detecting a deletion of an exception for the Firewall, sir"
		sText7 ="Sir, should I be deleting an exception to the Firewall?"
		sText8 ="the Firewall has had an exception deleted."
		sText9 ="Sir, are you deleting an exception to the Firewall?"
		sText0 ="the Firewall has an exception deleted"
		
		
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