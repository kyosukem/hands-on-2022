Option Explicit

'----------------------------------------------------------------------------------------
' 準備処理
'----------------------------------------------------------------------------------------

' 入力値チェック関数の読み込み
Include("functions\input_check.vbs")

' 定数定義の読み込み
Include("common\constant_definition.vbs")

'----------------------------------------------------------------------------------------
' main
'----------------------------------------------------------------------------------------

' ==============================
' 入力値１の読み込み＆チェック
' ==============================
dim varChecked1
varChecked1 = input_check( Inputbox("１つ目の数字を入力してください","入力１") )

' チェック結果の取り出し
If varChecked1(MESSAGE) = "" then
	
	' 入力値のエラーなし場合、整数値を変数に取り出す
	Dim intValue1
	intValue1 = varChecked1(VALUE)
	
else
	
	' 入力値のエラーの場合、エラーメッセージを表示して処理を終了する。
	msgbox("PGM終了:" & varChecked1(MESSAGE))
	WScript.Quit
	
end If

' ==============================
' 入力値２の読み込み＆チェック
' ==============================
dim varChecked2
varChecked2 = input_check( Inputbox("２つ目の数字を入力してください","入力２") )

' チェック結果の取り出し
If varChecked2(MESSAGE) = "" then
	
	' 入力値のエラーなし場合、整数値を変数に取り出す
	Dim intValue2
	intValue2 = varChecked2(VALUE)
	
else
	
	' 入力値のエラーの場合、エラーメッセージを表示して処理を終了する。
	msgbox("PGM終了:" & varChecked2(MESSAGE))
	WScript.Quit
	
end If

' ==============================
' 計算・結果表示
' ==============================
Dim intResult
intResult = intValue1 + intValue2

msgbox("計算結果：" & intResult)

'----------------------------------------------------------------------------------------
' end
'----------------------------------------------------------------------------------------


'----------------------------------------------------------------------------------------
' 関数
'----------------------------------------------------------------------------------------

' 外部ファイルの取り込み
Function Include(strFileName)
	
    Dim objFso
    Set objFso = Wscript.CreateObject("Scripting.FileSystemObject")
    
    Dim objWsh
    Set objWsh = objFso.OpenTextFile(strFileName,1)
	ExecuteGlobal objWsh.ReadAll()
	objWsh.Close
 
    Set objWsh = Nothing
    Set objFso = Nothing
    
End Function
