Dim speak, time

Set speak = CreateObject("sapi.spvoice")

time = InputBox("Enter time in seconds for alarm:")

WScript.Sleep time * 1000

speak.Speak "Wake up Implementation Engineer! It's time to code!"

MsgBox " Wake up Implementation Engineer! ", 0, "Alarm Clock"
