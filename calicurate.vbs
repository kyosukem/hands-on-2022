Option Explicit

Dim varInput1
varInput1 = Inputbox("１つ目の数字を入力してください","入力１")

If IsNumeric(varInput1) = False then
	msgbox("数値ではありません")
	WScript.Quit
end If


Dim varInput2
varInput2 = Inputbox("２つ目の数字を入力してください","入力２")

If IsNumeric(varInput2) = False then
	msgbox("数値ではありません")
	WScript.Quit
end If

Dim intResult
intResult = CInt(varInput1) + CInt(varInput2)

msgbox("計算結果：" & intResult)