Dim voice, fileStream, fso, timestamp, filePath, userText

' Prompt user for input text
userText = InputBox("Enter the text to convert to speech:", "Text to Speech")

' Exit if user cancels or leaves input empty
If Trim(userText) = "" Then
    WScript.Echo "No input provided. Exiting."
    WScript.Quit
End If

Set voice = CreateObject("SAPI.SpVoice")
Set fileStream = CreateObject("SAPI.SpFileStream")
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a safe timestamp string for filename
timestamp = Replace(Replace(Replace(Now, ":", "-"), "/", "-"), " ", "_")
filePath = fso.GetAbsolutePathName(".") & "\tts_" & timestamp & ".wav"

' Open stream for writing (3 = SSFMCreateForWrite)
fileStream.Open filePath, 3, False

' Assign the stream and speak the input text
Set voice.AudioOutputStream = fileStream
voice.Speak userText

' Close the stream
fileStream.Close

' Notify user of success
MsgBox "Speech exported to:" & vbCrLf & filePath, vbInformation, "Done"
