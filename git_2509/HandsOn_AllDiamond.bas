Attribute VB_Name = "HandsOn"
Option Explicit

Sub RightUp_01()
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 1 To i
            Cells(i, j).Value = "��"
        Next j
    Next i
    
'cells(1,1)�́�1�Acells(2,2)��2��...�ƍs(��)�ԍ��Ɠ���������
'A�񂩂珇�ɑ����鋓���ɂ���(1��2��3...)����j�̏������E�ӂ�i
           
End Sub

Sub LeftUp_02()
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 11 - i To 10
            Cells(i, j).Value = "��"
        Next j
    Next i
    
'�Ei���s�Aj�����S��
'�Ei��j��11�ɂȂ�悤����  cells(1, 10)��cells(2, 9)

'�E�uj��10�ɂȂ�܂Łv�Ƃ��������𖞂����A������͂���Z�����w�肷�邽��
'���̏����𖞂����ɂ�
'1.�s/��ԍ��̍ő�l(�����10)���傫�Ȑ�����For�̏�����


      
End Sub

Sub RightDown_03()
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 1 To 11 - i
            Cells(i, j).Value = "��"
        Next j
    Next i
    
End Sub


Sub LeftDown_04()
  Dim i As Long, j As Long
  For i = 1 To 10
      For j = i To 10
          Cells(i, j).Value = "��"
      Next j
  Next i
  
  
'1�s�ڂ͂��ׂā��A��������P�����炷�̂�i = 1 to 10
'i���P�������Ă����̂ŁA���̓��͂��J�n��������̗�ԍ��Ƃ���j�̏�������i����X�^�[�g(j = i to 10)
  
End Sub

Sub RightHalf_05()
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 1 To i
            Cells(i, j).Value = "��"
        Next j
    Next i
    
    For i = 11 To 19
        For j = 1 To 20 - i
            Cells(i, j).Value = "��"
        Next j
    Next i
    
'10�s�ڂ܂ł�RightUp_01�Ɠ���
'11�s�ڂ�A�񂩂灡9�A��������19�s�ڂ܂�1������̂ŊO����i = 11 to 19
'cells(11, 1)~cells(11, 9)�܂œ��͂�����&i���g���čs�������䂷�遨j��1����X�^�[�g
'�������̉E�ӂ�9�ɂ���Ǝl�p�`�Ɂ����o�͂���Ă��܂��̂ŁA�󔒂Ƃ������Z��(J11,I12...)�����[�v���d�˂閈�ɑ��₷���ߊO����i���E�ӂƂ���


   
End Sub


Sub LeftHalf_06()
    Dim i As Long, j As Long
    '�l�X�g1
    For i = 1 To 10
        For j = 11 - i To 10
            Cells(i, j).Value = "��"
        Next j
    Next i

    '�l�X�g2
    For i = 11 To 19
        For j = i - 9 To 10
            Cells(i, j).Value = "��"
        Next j
    Next i

    
'10�s�ڂ܂ł�LeftUp_02�Ɠ���

'�l�X�g2�̍s�𐧌䂷��i�́A�l�X�g1�܂łƑΏ̂Ƃ��邽��11 to 19
'�l�X�g2��B11(cells(11, 2))���灡����͂���������2�����ӁA10���E�ӂƂ�����
'�s�𐧌䂵�ē��͂��Ȃ���(11�s�ڂ�A�A12�s�ڂ�AB...)����邽�߁A���ӂ�1������悤i(11 to 19) - 9�Ƃ���

    
End Sub


Sub UpHalf_07()
'    Dim i As Long, j As Long
'    For i = 1 To 10
'        For j = 11 - i To 10
'            Cells(i, j).Value = "��"
'        Next j
'    Next i
'
'    For i = 1 To 10
'        For j = 10 + i To 19
'            Cells(i, j).Value = "��"
'        Next j
'    Next i


    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 11 - i To 10
            Cells(i, j).Value = "��"
        Next j
    Next i
    
    For i = 2 To 10
        For j = 11 To 9 + i
           Cells(i, j).Value = "��"
        Next j
    Next i
    
End Sub


Sub DouwHalf_08()
    Dim i As Long, j As Long
    
    For i = 1 To 10
        For j = 0 + i To 10
             Cells(i, j).Value = "��"
        Next j
    Next i
'j��0����X�^�[�g���邱�Ƃ�i(��ԍ�)�Ɠ����X�^�[�g�ʒu��
    
    For i = 1 To 10
        For j = 11 To 20 - i
            Cells(i, j).Value = "��"
        Next j
    Next i
'K��(Cells(1, 11))���������P�����炵����
'�̂ŁA���ӂ�11�A�E�ӂ�20 - i�Ƃ��ĕω�������
    
End Sub


Sub AALDiamond_09()
'����
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 11 - i To 10
            Cells(i, j).Value = "��"
        Next j
    Next i
    
'�E��
    For i = 2 To 10
        For j = 11 To 9 + i
            Cells(i, j).Value = "��"
        Next j
    Next i
    
'����
    For i = 11 To 20
        For j = i - 9 To 10
            Cells(i, j).Value = "��"
        Next j
    Next i
    
'�E��
    For i = 11 To 19
        For j = 11 To 29 - i
            Cells(i, j).Value = "��"
        Next j
    Next i
        
End Sub


Sub test()
    Dim i As Long, j As Long
'    For i = 1 To 10
'        Cells(i, 1) = "��"
'    Next i
    For j = 1 To 10
        Cells(1, j) = "��"
    Next j
    
End Sub
