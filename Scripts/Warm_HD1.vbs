
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
		
		zText0 ="Sir,I Am detecting a warm Hard Drive, drive 1 is above 70 degrees?"
		sText1 ="Sir,Hard Drive1 is above 70 degrees"
		sText2 ="Sir,Hard Drive #1 is above 70 degrees, getting a little warm"
		sText3 ="Sir,Hard Drive #1 is warm, above 70 degrees"
		sText4 ="Sir,Hard Drive #1 is above 70 degrees"
		sText5 ="Sir,You may want to check Hard Drive #1, the temperature is above 70 degrees."
		sText6 ="Sir,Hard Drive #1 temperatures are above 70 degrees "
		sText7 ="Sir, Would you please check Hard Drive #1's temperature? it is above 70 degrees"
		sText8 ="Sir,I think I may have a fever! Hard Drive 1's temperature is above 70 degrees"
		sText9 ="Sir, Is it hot in here to you? Hard Drive #1 is above 70 degrees"
		sText10 ="Sir,one of my drives is a little warm, Hard Drive #1 is above 70 degrees"
		sText11 ="Sir,a drive is warm, Hard Drive #1 is above 70 degrees"
		sText12 ="Sir,Drive #1 is warm,it is above 70 degrees"
		
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
