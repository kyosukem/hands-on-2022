Option Explicit

'----------------------------------------------------------------------------------------
' ���͒l�̃`�F�b�N���[�`��
'
'   ���́ivarInput�j     �F�`�F�b�N�Ώۂ̕ϐ��B�o���A���g�^�B
'
'   �ԋp�l�iinput_check�j�F�����̒l��ԋp���邽�߁A�Q�v�f�̔z��Ƃ���B
'
'                �P�ڂ̗v�f�F�G���[���b�Z�[�W
'                    �`�F�b�N���ʂŃG���[�Ȃ��F�����[���̕�����i""�j
'                    �`�F�b�N���ʂŃG���[����F���b�Z�[�W������
'
'                �Q�ڂ̗v�f�F�ϊ��㐔�l
'                    �`�F�b�N���ʂŃG���[�Ȃ��FvarInput�𐮐��ϕϊ������l�BInteger�^�B
'                    �`�F�b�N���ʂŃG���[����FvbNull
'----------------------------------------------------------------------------------------
Function input_check (varInput)
	
	' �G���[�F���͒l�����l�łȂ��ꍇ
	If IsNumeric(varInput) = False then
		input_check = array("���l�ł͂���܂���",vbNull)
		Exit Function
	end If
	
	' ���폈���F���͒l�𐮐��ɕϊ����ĕԋp����
	input_check = array("",CInt(varInput))
	
End Function
