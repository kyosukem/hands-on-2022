Option Explicit

'----------------------------------------------------------------------------------------
' 単体テスト実行メイン
'
' USEAGE:
' コマンドプロンプトから、以下のコマンドで実行する。
' （WScript.StdOutによる標準出力への書き出しは、cscript配下でないと実施できないため）
' > cd ＜インストールされたフォルダ＞
' > cscript calicurate_test.vbs
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' 準備処理
'----------------------------------------------------------------------------------------

' 入力値チェック関数の読み込み
Include("..\functions\input_check.vbs")

' 定数定義の読み込み
Include("..\common\constant_definition.vbs")

'----------------------------------------------------------------------------------------
' main
'----------------------------------------------------------------------------------------

Dim intResult
intResult = 0

WScript.StdOut.Write("テストケース１：ノーマルケース　引数がプラスの整数" & vbCRLF)
intResult = intResult + test_execute( "12345" , 12345 , "")

WScript.StdOut.Write("テストケース２：ノーマルケース　引数がマイナスの整数" & vbCRLF)
intResult = intResult + test_execute( "-12345" , -12345 , "")

WScript.StdOut.Write("テストケース３：ノーマルケース　引数がゼロ" & vbCRLF)
intResult = intResult + test_execute( "0" , 0 , "")

WScript.StdOut.Write("テストケース４：エラーケース　引数が数値でない（アルファベット）" & vbCRLF)
intResult = intResult + test_execute( "a" , vbNull , "数値ではありません")

WScript.StdOut.Write("テストケース５：エラーケース　引数が数値でない（記号）" & vbCRLF)
intResult = intResult + test_execute( "." , vbNull , "数値ではありません")

WScript.StdOut.Write("テストケース６：エラーケース　引数が数値でない（数字と記号が混在）" & vbCRLF)
intResult = intResult + test_execute( "1<" , vbNull , "数値ではありません")


' 小数点ありはissue6にてエラーにする予定。現時点ではノーマルケースとして扱う
WScript.StdOut.Write("テストケース７：ノーマルケース　引数が数値でない（数字とピリオドが混在:前）" & vbCRLF)
intResult = intResult + test_execute( ".1" , 0 , "")
WScript.StdOut.Write("テストケース８：ノーマルケース　引数が数値でない（数字とピリオドが混在:後）" & vbCRLF)
intResult = intResult + test_execute( "1." , 1 , "")


' 全角文字ありはissue11にて議論中。現時点では、int関数による自動変換によるOKケースとする。
WScript.StdOut.Write("テストケース１０：ノーマルケース　引数が全角の数字" & vbCRLF)
intResult = intResult + test_execute( "１０" , 10 , "")
WScript.StdOut.Write("テストケース１１：ノーマルケース　引数が半角・全角の数字混在" & vbCRLF)
intResult = intResult + test_execute( "2１０" , 210 , "")



' テスト成否判定（intResultがゼロであれば全テストケース成功）
If intResult = 0 then
	WScript.StdOut.Write("---------------------" & vbCRLF)
	WScript.StdOut.Write("  TEST SUCCESS       " & vbCRLF)
	WScript.StdOut.Write("---------------------" & vbCRLF)
else
	WScript.StdOut.Write("!!!!!!!!!!!!!!!!!!!!!" & vbCRLF)
	WScript.StdOut.Write("  TEST FAILURE       " & vbCRLF)
	WScript.StdOut.Write("!!!!!!!!!!!!!!!!!!!!!" & vbCRLF)
End If

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

'----------------------------------------------------------------------------------------
' input_checkのテストルーチン
'   input_checkの入力値、期待値を受け取り、input_checkの実行結果と比較する。
'
'   入力（varArg）             ：チェック対象の変数。バリアント型。
'   入力（expectation_value）  ：VALUEに対する期待値
'   入力（expectation_message）：MESSAGEに対する期待値
'
'   返却値（test_execute）：
'                    チェック結果でエラーなし：0
'                    チェック結果でエラーあり：1
'----------------------------------------------------------------------------------------

Function test_execute(varArg , expectation_value , expectation_message)
	
	WScript.StdOut.Write("引数：" & varArg & vbCRLF)
	
	dim varChecked
	varChecked = input_check(varArg)
	WScript.StdOut.Write("VALUE:" & varChecked(VALUE) & "  MESSAGE:" & varChecked(MESSAGE) & vbCRLF)
	
	If (varChecked(VALUE)   = expectation_value  ) and _
	   (varChecked(MESSAGE) = expectation_message) then
		
		WScript.StdOut.Write("テスト成功" & vbCRLF)
		test_execute = 0
		
	else
		
		WScript.StdOut.Write("テスト失敗" & vbCRLF)
		test_execute = 1
		
	end If
	
	WScript.StdOut.Write(vbCRLF)
	
End Function
