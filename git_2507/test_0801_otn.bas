Attribute VB_Name = "test_0801_otn"
Option Explicit

Sub mcr01()
'  A1�Z���Ɂu����ɂ��́v�ƕ\��������}�N�����쐬���Ȃ����B
  Cells(1, 1).Value = "����ɂ���"
End Sub

Sub mcr02()
'  Cells���g���āAB3�Z���ɁuVBA�e�X�g�v�ƕ\�������Ȃ����B
  Cells(3, 2).Value = "VBA�e�X�g"
End Sub


Sub mcr03()
'  ���l�^�̕ϐ�price��錾���āu1200�v�������AA5�Z���ɏo�͂�����}�N�����쐬���Ȃ����B
  Dim price As Long
  price = 1200
  Cells(5, 1).Value = price
End Sub

Sub mcr04()
'  A1�Z���Ɂu�����v�AB1�Z���Ɂu�_���v�Ɠ��͂���}�N�����쐬���Ȃ����B
  Cells(1, 1).Value = "����"
  Cells(2, 1).Value = "�_��"
  
End Sub


Sub mcr05()
  '���l�u2000�v��萔�uPrice�v�Ƃ��Đ錾���AA2�Z���ɕ\������}�N�����쐬���Ȃ����B
  Const price As Long = 2000
  Cells(2, 1).Value = price
  
End Sub

Sub mcr06()
'  �ϐ�greeting�Ɂu���͂悤�v�������AB5�Z���ɏo�͂���}�N�����쐬���Ȃ����B
  Dim greeting As String
  greeting = "���͂悤"
  Cells(5, 2).Value = greeting
  
End Sub

Sub mcr07()
'  �ϐ����o�R���ăZ���ɒl��\�������{�I�ȃ}�N�����쐬���Ȃ����
  Dim today As String
  today = "8��1��"
  Cells(1, 3) = today
  
End Sub


Sub mcr08()
 '�ϐ�score��85�������ACells���g����B5�Z���ɏo�͂���}�N�����쐬���Ȃ����B
 Dim score As Long
 score = 85
 Cells(5, 2).Value = score
 
End Sub
Sub mcr09()
'  �ϐ�firstName�Ɂu�R�c�v�AlastName�Ɂu���Y�v�������A��������������A1�Z���ɏo�͂���}�N�����쐬���Ȃ����B
  Dim firstName As String, lastName As String
  firstName = "�R�c"
  lastName = "���Y"
  Cells(1, 1).Value = firstName & lastName
  
End Sub

Sub mcr10()
 '�萔�uBASIC�v��80�A�ϐ��uAdd�v��20�Ƃ��āA���v�_���v�Z��C1�Z���ɕ\������}�N�����쐬���Ȃ����B
 Const BASIC As Long = 80
 Dim add As Long
 add = 20
 Cells(1, 3).Value = BASIC + add
End Sub
