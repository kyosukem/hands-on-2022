Option Explicit

Dim varInput1
varInput1 = Inputbox("‚P‚Â–Ú‚Ì”š‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢","“ü—Í‚P")

If IsNumeric(varInput1) = False then
	msgbox("”’l‚Å‚Í‚ ‚è‚Ü‚¹‚ñ")
	WScript.Quit
end If


Dim varInput2
varInput2 = Inputbox("‚Q‚Â–Ú‚Ì”š‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢","“ü—Í‚Q")

If IsNumeric(varInput2) = False then
	msgbox("”’l‚Å‚Í‚ ‚è‚Ü‚¹‚ñ")
	WScript.Quit
end If

Dim intResult
intResult = CInt(varInput1) + CInt(varInput2)

msgbox("ŒvZŒ‹‰ÊF" & intResult)