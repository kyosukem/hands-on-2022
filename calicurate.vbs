Option Explicit

Dim varInput1
varInput1 = Inputbox("�P�ڂ̐�������͂��Ă�������","���͂P")

If IsNumeric(varInput1) = False then
	msgbox("���l�ł͂���܂���")
	WScript.Quit
end If


Dim varInput2
varInput2 = Inputbox("�Q�ڂ̐�������͂��Ă�������","���͂Q")

If IsNumeric(varInput2) = False then
	msgbox("���l�ł͂���܂���")
	WScript.Quit
end If

Dim intResult
intResult = CInt(varInput1) + CInt(varInput2)

msgbox("�v�Z���ʁF" & intResult)