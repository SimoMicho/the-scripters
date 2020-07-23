
intAnswer = _
    Msgbox("Do you want a male voice?", _
        vbYesNo, "Notepad TTS Options")
If intAnswer = vbYes Then
dim message , reader
message = inputbox("Please, insert a message.","Notepad Text To Speech")
set reader = createobject("sapi.spvoice")
reader.speak message

Else
    dim msg, sapi, result
result = 6

set sapi = CreateObject("sapi.spvoice")
set sapi.Voice = sapi.GetVoices.Item(1)
sapi.rate=0
sapi.volume = 90


do while result = 6
msg = InputBox("Please, insert a message.","Notepad Text To Speech")
sapi.Speak msg
result = msgbox("Would you like to do it again?",4,"Notepad Text To Speech")
loop
End If