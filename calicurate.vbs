Option Explicit

Dim varInput1
varInput1 = Inputbox("�P�ڂ̐�������͂��Ă�������","���͂P")

Dim varInput2
varInput2 = Inputbox("�Q�ڂ̐�������͂��Ă�������","���͂Q")

Dim intResult
intResult = CInt(varInput1) + CInt(varInput2)

msgbox("�v�Z���ʁF" & intResult)