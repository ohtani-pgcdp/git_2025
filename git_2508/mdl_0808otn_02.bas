Attribute VB_Name = "mdl_0808otn_02"
Option Explicit


'*********************************************************
'* ForNext�̗��K
'*********************************************************

Sub mcr0808_01()
  Dim i               '�ϐ�i��錾
  For i = 1 To 10    '�ϐ�i���P����10�܂łP���J�E���g�A�b�v
    Cells(i, 1) = 1  'A���i�s�ڂ̃Z����1������
  Next i               'For�͈̔͂͂����܂�


End Sub

Sub mcr0808_02()
  Dim i
  For i = 10 To 1 Step -1   '�ϐ�i���P����10�܂łP���J�E���g�A�b�v
    Cells(i, 3) = 1  'A���i�s�ڂ̃Z����1������
  Next i
End Sub

Sub mcr0808_03()
 Dim i
  For i = 10 To 1 Step -1   '�ϐ�i���P����10�܂łP���J�E���g�A�b�v
    Cells(i, 7) = 1  'A���i�s�ڂ̃Z����1������
  Next i
  
End Sub


Sub mcr0808_05()
  Dim i
    For i = 1 To 10
        If Cells(i, 9) <> "" Then
'            Exit For
        End If
        Cells(i, 9) = 1
    Next
    
'�u<>�v�́A�u�����ł͂Ȃ��v�Ƃ������Ƃ𒲂ׂ鉉�Z�q�ł��B
'�u<>�v�̍��ƉE���ׂāA�������Ȃ����Ƃ𒲂ׂ��r���Z�q�ł��B



End Sub

' ���ׂ�F�C���N�������g�͂ǂ��ōs���Ă�H

Sub mcr0808_06()
  Dim i, j
      For i = 15 To 25
          For j = 1 To 10
              Cells(i, j) = 1
          Next j
      Next i

End Sub

'�P���
'�ȉ����R��A���b�Z�[�W�{�b�N�X�ɕ\��������B
'�u�T�T�v���R��\��

Sub mcr0808_07()
  Dim i
    For i = 1 To 3
      MsgBox 55
    Next i
    
End Sub

'�Q���
'�ȉ������b�Z�[�W�{�b�N�X�ɕ\��
'���Ȃ���@�Ƃ�1��ڂƂ����Ӗ�
'�@5
'�A10
'�B15

Sub mcr0808_08()
  Dim i
    For i = 5 To 15 Step 5
      MsgBox i
    Next i
    
End Sub


