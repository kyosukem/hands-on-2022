Option Explicit

'----------------------------------------------------------------------------------------
' �P�̃e�X�g���s���C��
'
' USEAGE:
' �R�}���h�v�����v�g����A�ȉ��̃R�}���h�Ŏ��s����B
' �iWScript.StdOut�ɂ��W���o�͂ւ̏����o���́Acscript�z���łȂ��Ǝ��{�ł��Ȃ����߁j
' > cd ���C���X�g�[�����ꂽ�t�H���_��
' > cscript calicurate_test.vbs
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' ��������
'----------------------------------------------------------------------------------------

' ���͒l�`�F�b�N�֐��̓ǂݍ���
Include("..\functions\input_check.vbs")

' �萔��`�̓ǂݍ���
Include("..\common\constant_definition.vbs")

'----------------------------------------------------------------------------------------
' main
'----------------------------------------------------------------------------------------

Dim intResult
intResult = 0

WScript.StdOut.Write("�e�X�g�P�[�X�P�F�m�[�}���P�[�X�@�������v���X�̐���" & vbCRLF)
intResult = intResult + test_execute( "12345" , 12345 , "")

WScript.StdOut.Write("�e�X�g�P�[�X�Q�F�m�[�}���P�[�X�@�������}�C�i�X�̐���" & vbCRLF)
intResult = intResult + test_execute( "-12345" , -12345 , "")

WScript.StdOut.Write("�e�X�g�P�[�X�R�F�m�[�}���P�[�X�@�������[��" & vbCRLF)
intResult = intResult + test_execute( "0" , 0 , "")

WScript.StdOut.Write("�e�X�g�P�[�X�S�F�G���[�P�[�X�@���������l�łȂ��i�A���t�@�x�b�g�j" & vbCRLF)
intResult = intResult + test_execute( "a" , vbNull , "���l�ł͂���܂���")

WScript.StdOut.Write("�e�X�g�P�[�X�T�F�G���[�P�[�X�@���������l�łȂ��i�L���j" & vbCRLF)
intResult = intResult + test_execute( "." , vbNull , "���l�ł͂���܂���")

WScript.StdOut.Write("�e�X�g�P�[�X�U�F�G���[�P�[�X�@���������l�łȂ��i�����ƋL�������݁j" & vbCRLF)
intResult = intResult + test_execute( "1<" , vbNull , "���l�ł͂���܂���")


' �����_�����issue6�ɂăG���[�ɂ���\��B�����_�ł̓m�[�}���P�[�X�Ƃ��Ĉ���
WScript.StdOut.Write("�e�X�g�P�[�X�V�F�m�[�}���P�[�X�@���������l�łȂ��i�����ƃs���I�h������:�O�j" & vbCRLF)
intResult = intResult + test_execute( ".1" , 0 , "")
WScript.StdOut.Write("�e�X�g�P�[�X�W�F�m�[�}���P�[�X�@���������l�łȂ��i�����ƃs���I�h������:��j" & vbCRLF)
intResult = intResult + test_execute( "1." , 1 , "")


' �S�p���������issue11�ɂċc�_���B�����_�ł́Aint�֐��ɂ�鎩���ϊ��ɂ��OK�P�[�X�Ƃ���B
WScript.StdOut.Write("�e�X�g�P�[�X�P�O�F�m�[�}���P�[�X�@�������S�p�̐���" & vbCRLF)
intResult = intResult + test_execute( "�P�O" , 10 , "")
WScript.StdOut.Write("�e�X�g�P�[�X�P�P�F�m�[�}���P�[�X�@���������p�E�S�p�̐�������" & vbCRLF)
intResult = intResult + test_execute( "2�P�O" , 210 , "")



' �e�X�g���۔���iintResult���[���ł���ΑS�e�X�g�P�[�X�����j
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

'----------------------------------------------------------------------------------------
' input_check�̃e�X�g���[�`��
'   input_check�̓��͒l�A���Ғl���󂯎��Ainput_check�̎��s���ʂƔ�r����B
'
'   ���́ivarArg�j             �F�`�F�b�N�Ώۂ̕ϐ��B�o���A���g�^�B
'   ���́iexpectation_value�j  �FVALUE�ɑ΂�����Ғl
'   ���́iexpectation_message�j�FMESSAGE�ɑ΂�����Ғl
'
'   �ԋp�l�itest_execute�j�F
'                    �`�F�b�N���ʂŃG���[�Ȃ��F0
'                    �`�F�b�N���ʂŃG���[����F1
'----------------------------------------------------------------------------------------

Function test_execute(varArg , expectation_value , expectation_message)
	
	WScript.StdOut.Write("�����F" & varArg & vbCRLF)
	
	dim varChecked
	varChecked = input_check(varArg)
	WScript.StdOut.Write("VALUE:" & varChecked(VALUE) & "  MESSAGE:" & varChecked(MESSAGE) & vbCRLF)
	
	If (varChecked(VALUE)   = expectation_value  ) and _
	   (varChecked(MESSAGE) = expectation_message) then
		
		WScript.StdOut.Write("�e�X�g����" & vbCRLF)
		test_execute = 0
		
	else
		
		WScript.StdOut.Write("�e�X�g���s" & vbCRLF)
		test_execute = 1
		
	end If
	
	WScript.StdOut.Write(vbCRLF)
	
End Function
