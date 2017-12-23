
Const SVSFlagsAsync = 1

Dim Speech
Dim FSO
Dim max,min
max=12
min=0
Randomize
CreateObjects
Main
DestroyObjects
Quit

Sub Main
        Dim sText0
        Dim sText1
		Dim sText2
		Dim sText3
		Dim sText4
		Dim sText5
		Dim sText6
		Dim sText7
		Dim sText8
		Dim sText9
		Dim sText10
		Dim sText11
		Dim sText12
		
		zText0 ="Sir,I Am detecting a problem with the water cooling pump, it has stopped working."
		sText1 ="Sir,the water cooling pump has a problem"
		sText3 ="Sir,the water cooling pump has stopped working"
		sText4 ="Sir, would you please check the water cooling pump"
		sText5 ="Sir,You may want to check the water cooling pump."
		sText6 ="Sir,the water cooling pump has stopped "
		sText7 ="Sir, Would you please check the water cooling pump?"
		sText8 ="Sir,I think the water cooling pump has stopped working!"
		sText9 ="Sir, Is it hot in here to you? the water cooling pump has stopped working"
		sText10 ="Sir,I am operating a little warm, the water cooling pump has stopped working"
		sText11 ="Sir,I am a little warm, the water cooling pump has stopped working"
		sText12 ="Sir,processors Temperatures are warm,the water cooling pump has stopped working"
		
       sText = (Int((max-min)*Rnd+min))
        sText = Trim(sText)
         If sText = "0" Then
                SpeakText sText0
				else
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
		 If sText = "10" Then
                SpeakText sText10
				else
		If sText = "11" Then
                SpeakText sText11
				else
		If sText = "12"  Then
		        SpeakText sText12
				
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
