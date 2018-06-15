Dim message,orton
message = InputBox("This is for you"+vbcrlf+"Enjoy","Think different")
Set orton = CreateObject("sapi.spvoice")
orton.Speak message