Attribute VB_Name = "test_0801_otn"
Option Explicit

Sub mcr01()
'  A1�Z���Ɂu����ɂ��́v�ƕ\��������}�N�����쐬���Ȃ����B
  Cells(1, 1).Value = "����ɂ���"
End Sub


Sub mcr03()
'  ���l�^�̕ϐ�price��錾���āu1200�v�������AA5�Z���ɏo�͂�����}�N�����쐬���Ȃ����B
  Dim price As Long
  price = 1200
  Cells(5, 1).Value = price
End Sub

Sub mcr05()
  '���l�u2000�v��萔�uPrice�v�Ƃ��Đ錾���AA2�Z���ɕ\������}�N�����쐬���Ȃ����B
  Const price As Long = 2000
  Cells(2, 1).Value = price
  
End Sub


Sub mcr08()
 '�ϐ�score��85�������ACells���g����B5�Z���ɏo�͂���}�N�����쐬���Ȃ����B
 Dim score As Long
 score = 85
 Cells(5, 2) = score
 
End Sub


Sub mcr10()
 '�萔�uBASIC�v��80�A�ϐ��uAdd�v��20�Ƃ��āA���v�_���v�Z��C1�Z���ɕ\������}�N�����쐬���Ȃ����B
 
End Sub
