Dim say,speech
say="Welcome to Windows 10!! Happy programming"
set speech=CreateObject("SAPI.SpVoice")
speech.Speak say
