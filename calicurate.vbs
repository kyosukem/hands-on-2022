Option Explicit

'----------------------------------------------------------------------------------------
' ��������
'----------------------------------------------------------------------------------------

' ���͒l�`�F�b�N�֐��̓ǂݍ���
Include("functions\input_check.vbs")

' �萔��`�̓ǂݍ���
Include("common\constant_definition.vbs")

'----------------------------------------------------------------------------------------
' main
'----------------------------------------------------------------------------------------

' ==============================
' ���͒l�P�̓ǂݍ��݁��`�F�b�N
' ==============================
dim varChecked1
varChecked1 = input_check( Inputbox("�P�ڂ̐�������͂��Ă�������","���͂P") )

' �`�F�b�N���ʂ̎��o��
If varChecked1(MESSAGE) = "" then
	
	' ���͒l�̃G���[�Ȃ��ꍇ�A�����l��ϐ��Ɏ��o��
	Dim intValue1
	intValue1 = varChecked1(VALUE)
	
else
	
	' ���͒l�̃G���[�̏ꍇ�A�G���[���b�Z�[�W��\�����ď������I������B
	msgbox("PGM�I��:" & varChecked1(MESSAGE))
	WScript.Quit
	
end If

' ==============================
' ���͒l�Q�̓ǂݍ��݁��`�F�b�N
' ==============================
dim varChecked2
varChecked2 = input_check( Inputbox("�Q�ڂ̐�������͂��Ă�������","���͂Q") )

' �`�F�b�N���ʂ̎��o��
If varChecked2(MESSAGE) = "" then
	
	' ���͒l�̃G���[�Ȃ��ꍇ�A�����l��ϐ��Ɏ��o��
	Dim intValue2
	intValue2 = varChecked2(VALUE)
	
else
	
	' ���͒l�̃G���[�̏ꍇ�A�G���[���b�Z�[�W��\�����ď������I������B
	msgbox("PGM�I��:" & varChecked2(MESSAGE))
	WScript.Quit
	
end If

' ==============================
' �v�Z�E���ʕ\��
' ==============================
Dim intResult
intResult = intValue1 + intValue2

msgbox("�v�Z���ʁF" & intResult)

'----------------------------------------------------------------------------------------
' end
'----------------------------------------------------------------------------------------


'----------------------------------------------------------------------------------------
' �֐�
'----------------------------------------------------------------------------------------

' �O���t�@�C���̎�荞��
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
