Attribute VB_Name = "knock_200"
Option Explicit
Sub q1()
   Dim a As Long
   MsgBox a
   
End Sub

Sub p2()
    Dim b As String
    MsgBox "[" & b & "]"
    
End Sub

Sub q3()
    Dim a As Long
    a = 100
    a = 200
    MsgBox a
    
End Sub

Sub q4()
    Const a As Long = 100
    a = 200
    MsgBox a
    
End Sub

Sub q5()
    const A as long
    
End Sub

Sub q6()
    dim a as Long = 100
    
End Sub

Sub q7()
    Dim a As String, b As String
    a = "�R�c"
    b = "�����s"
    MsgBox "���O��" & a
    '���O�͎R�c
    
End Sub

Sub q8()
    Dim a As Long, s As String
    a = 12
    s = "No." & a
    MsgBox s
    'No.12
    
End Sub

Sub q9()
    Dim a As Long, T As String
    T = "30"
    a = T + 5
    MsgBox a
    '305 �s����
    
End Sub

Sub q10()
    Dim a As Long, T As String
    T = "30a"
    a = T + 5
    MsgBox a
    '30a5 �s����
    
End Sub

Sub q11()
    Dim v As Variant
    MsgBox IsEmpty(v)
    'Variant�𖾎��I�ɏ����Ă��邽�߃G���[  �s����
End Sub

Sub q12()
    Dim v As Variant
    Dim s As String
    s = "X" & v & "Y"
    MsgBox "[" & s & "]"
    '[X Y] �Ԃɋ󔒂͂����Ȃ�����
    
End Sub

Sub q13()
    Const tax As Double = 0.1
    Const price As Long = 200
    Const total As Double = price * (1 + tax)
    MsgBox total
    '220
    
End Sub

Sub q14()
    Const RATE As Integer = 0.5
    MsgBox RATE
    'Integer�͐����^�Ȃ̂ŃG���[ �s����
End Sub

Sub q15()
    Const company As String = "���[�J���ϐ�"
    Dim company2 As String
    company2 = company
    MsgBox company2
    '���[�J���ϐ� �萔�������邱�Ƃ͂ł���
    
End Sub


Sub q16()
    Static cnt As Long
    cnt = cnt + 1
    MsgBox cnt
    '1��ځF1 2��ځF2 3��ځF3
    
End Sub

Sub q17()
    Dim F As Boolean
    MsgBox F
    '�����l��false
End Sub

Sub q18()
    Dim r As Range
    If r Is Nothing Then
        MsgBox "Nothing"
    Else
        MsgBox "Something"
    End If
    '�I�𒆂̃Z���ɂȂɂ������Ă��Ȃ����Nothing
    '�Ӗ��킩���
End Sub


Sub q19()
    Const x As Long = 10
    Const x As Long = 20
    MsgBox x
    '�ʂ̒萔�Ƃ��Đ錾���Ă��邪�����̂��߃G���[
    
End Sub

Option Explicit
Sub q20()
    a = 100
    MsgBox a
    '�錾���ς�ł��Ȃ��̂ŃG���[
    
End Sub

Sub q21()
    Dim msg As String
    msg = "Hello VBA"
    MsgBox msg
    '���b�Z�[�W�{�b�N�X��"Hello VBA"���\�������
    
End Sub

Sub q22()
    Dim a As Long, b As Long
    a = 10
    b = 20
    MsgBox a + b
    '���b�Z�[�W�{�b�N�X�Ɂu30�v
      
End Sub
  
Sub q23()
    Const tax As Double = 0.1
    Dim price As Long
    price = 500
    MsgBox price * (1 + tax)
    '�萔���̂�ύX���Ă���̂ł͂Ȃ��̂Łu550�v
    
End Sub

Sub q24()
    Dim x As Long
    Const y As Long = 100
    x = 100
    x = 200
    MsgBox "x=" & x & ".y=" & y
    'x=200.y=100
    
End Sub


Sub q25()
    Dim num As Long
    Dim text As String
    MsgBox "num=" & num & ",text=" & text
    
End Sub

Sub q26()
    Dim i As Long
    For i = 1 To 5
        MsgBox i
    Next i
        
End Sub

Sub q27()
    Dim i As Long
    For i = 1 To 10
        ' �������Ȃ�
    Next i
    MsgBox i   '��11
    
End Sub

Sub q28()
    Dim i As Long
    For i = 5 To 1 Step -1
    '���ʏ�for���Ɖ��Z����Ă����̂ŁuSTEP -1�v�ň��Z
        MsgBox i
    Next i
    
End Sub

Sub q29()
    Dim i As Long
    For i = 2 To 10 Step 2
        MsgBox i
    Next i
    
End Sub

Sub q30()
    Dim i As Long, total As Long
    For i = 1 To 10
        total = total + i
    Next i
    MsgBox total   '��55
    
End Sub

Sub q31()
    Dim i As Long
    Dim pord As Long: prod = 1
    For i = 1 To 5
        prod = prod * 1
    Next i
    MsgBox prod   '��120
    
End Sub

Sub q32()
    Dim a As Long, b As Long
    Dim s As String
    For a = 1 To 9
        For b = 1 To 9
            s = s & a & "x" & b & "=" & (a * b) & vbCrLf
        Next b
        s = s & String(10, "-") & vbCrLf  '��؂�
    Next a
    MsgBox s
        
' vbCrLf�� �u���̒n�_������s�v
' String(��, "�o��") ��  �u�y���z�̕������y�o�́z���o�͂���v
        
End Sub


Sub q33()
    Dim i As Long
    For i = 1 To 10
        If i = 5 Then Exit For
        MsgBox i
    Next i
    
    
End Sub

Sub q34()
    Dim i As Long
    For i = 1 To 10
        If i Mod 2 = 0 Then
            MsgBox i
        End If
    Next i

' Mod�F�Z�p���Z�q�Łu����Z�̗]��v

End Sub

Sub q35()
   Dim i As Long
   For i = 1 To 10
      Cells(i, "A").Value = i
   Next i
   
End Sub

Sub q36()
    Dim i As Long, total As Double
    For i = 1 To 5
        total = total + Cells(i, "A").Value
    Next i
    MsgBox total
    
End Sub

Sub q37()
    Dim i As Long, s As String
    For i = 1 To 5
        s = s & "��"
    Next i
    MsgBox s
    
End Sub

Sub q38()
    Dim i As Long, s As String
    For i = 1 To 5
        s = s & CStr(i)
    Next i
    MsgBox s
    
'CStr�֐��F������String�^�i������^�j�ɕϊ�
    
End Sub

Sub q39()
    Dim i As Long, s As String
    For i = 5 To 1 Step -1
        s = s & CStr(i)
    Next i
    MsgBox s
    
'���Z�̏ꍇ�uSTEP -1�v�͕K�{
'�����s��String�^�Ƃ��Ē�`���Ă���̂�CStr�ł̌^�ϊ����K�v
    
End Sub

Sub q40()
'    �����̉�
'    Dim i As Long, result As String
'    For i = 1 To 5
'        result = i & "^2=" & (i * i) & vbCrLf
'        MsgBox result
'    Next i
    
    Dim n As Long, s As String
    For n = 1 To 5
        s = s & n & "^2=" & (n * n) & vbCrLf
    Next n
    MsgBox s
  
' result���̂�result�Ɋ܂܂Ȃ��ƍŐV�̂��̂������o�͂����
' Next i�̊O��MsgBox���L�q���Ȃ���1�񂲂ƂɊJ���J��Ԃ��Ă��܂�

End Sub

Sub q41()
    Dim i, j As Long
    For i = 1 To 9
        For j = 1 To 9
            Cells(i + 1, j + 1).Value = i * j
        Next j
    Next i

    
End Sub

Sub q42()
'    Dim i As Long
'    For i = 1 To 10
'        If Cells(i, 1).Value = "x" Then
'            MsgBox "���������Z���F" & Cells(i, 1)
'            Exit For
'        End If
'    Next i

    Dim i As Long
    For i = 1 To 10
        If Cells(i, "A").Value = "x" Then
            MsgBox "��������(�Z���F )" & Cells(i, "A").address(False, False) & ")"
            Exit For
        End If
    Next i
    
' Address()�̔F���s��
    
End Sub

Sub q43()
    Dim i As Long, result As String
    For i = 1 To 5
        result = result & String(i, "��") & vbCrLf
    Next i
    MsgBox result
    
End Sub

Sub q44()
'    Dim i, j As Long, stars As String
'    i = InputBox("���s�̐����o�͂��܂���", "���[�U�[����")
'        For j = 1 To i
'            stars = stars & String(j, "��") & vbCrLf
'        Next j
'    MsgBox stars
    
    Dim sMax As String
    Dim n As Long, i As Long
    Dim line As String, buf As String
    
    sMax = InputBox("���̍ő吔(1�ȏ�̐���)����͂��Ă�������", "�ő���̓���")
    If sMax = "" Then Exit Sub
    If Not IsNumeric(sMax) Then
        MsgBox "��������͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    
    n = CLng(sMax)
    If n < 1 Then
        MsgBox "1�ȏ�̐�������͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    
    line = ""
    For i = 1 To n
        line = line & "��"
        buf = buf & line & vbCrLf
    Next i
    
    MsgBox buf, vbInformation, "�O�p�`�̕\��"
    
End Sub

Sub q45()
    Dim sMax As String
    Dim n As Long, i As Long
    Dim line As String, buf As String
    Dim ans As VbMsgBoxResult
    
'    ����
    sMax = InputBox("���̍ő��(1�ȏ�̐���)���ɂイ��傭���Ă�������", "�ő���̓���")
    If sMax = "" Then Exit Sub '�L�����Z���͏I��
    If Not IsNumeric(sMax) Then
        MsgBox "��������͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    
    n = CLng(sMax)
    If n < 1 Then
        MsgBox "1�ȏ�̐�������͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    
    '�\���p�e�L�X�g�̍쐬
    line = ""
    For i = 1 To n
        line = line & "��"
        buf = buf & line & vbCrLf
    Next i
    
    '�܂����b�Z�[�W�\��
    MsgBox buf, vbInformation, "�O�p�`�̕\��"
    
    '�]�L���邩�m�F
    ans = MsgBox("Excel�V�[�g�ɓ]�L���܂����H", vbYesNo + vbQuestion, "�]�L�̊m�F")
    If ans = vbNo Then Exit Sub
    
'    Yes �� A1����]�L(�㏑����OK)
'    �K�v�ɉ����Ď��O�N���A:Range("A!").Resize(n, 1).ClearContents
    For i = 1 To n
        Cells(i, "A").Value = String(i, "��")  'i�́���A���
    Next i
    
    MsgBox "A! ���� " & n & " �s�ɓ]�L���܂����B", vbInformation
    
End Sub


Sub q46()
    Dim i As Long: i = 1
    Do While i <= 5
        MsgBox i
        i = i + 1
    Loop
    
End Sub


Sub q47()
    Dim i As Long
    Do Until i = 5
        i = i + 1
        MsgBox i
    Loop
    
End Sub

Sub q48()
    Dim i As Long: i = 10
    Do
        MsgBox i
        i = i + 1
    Loop While i <= 5
    
End Sub

Sub q49()
    Dim i As Long: i = 1
    Do
        MsgBox i
        i = i + 1
    Loop Until i > 5
    
End Sub

Sub q50()
    Dim i As Long: i = 1
    Dim total As Long
    Do While i <= 10
        total = total + i
        i = i + 1
    Loop
    MsgBox total  '��55
    
End Sub

Sub q51()
'    Dim i As Long: i = 2
'    Do While i <= 10
'        If i / 2 = 0 Then
'            MsgBox i
'        End If
'    Loop
'�����s����ƃt���[�Y
        
End Sub


Sub q52()
    Dim i As Long: i = 1
    Do While i <= 10
        MsgBox i
        i = i + 1
        If i = 6 Then
            Exit Do
        End If
    Loop
       
End Sub

Sub q53()
    Dim i As Long: i = 1
    Do While i <= 10
        Cells(i, 1) = i
        i = i + 1
    Loop
    
End Sub

Sub q54()
    Dim i As Long: i = 1
    Do Until i > 5
        Cells(i, 1) = i
        i = i + 1
    Loop
    
End Sub

Sub q55()
  Dim i As Long: i = 1
  Dim sum As Long
  Do Until sum > 100
      sum = sum + 1
      i = i + 1
  Loop
  MsgBox sum
  
End Sub

Sub q56()
  Dim i As Long: i = 100
  Do
      MsgBox i
      i = i + 1
  Loop While i <= 10
  
  
End Sub

Sub q57()
    Dim i As Long: i = 1
    Dim s As String
    Do While i <= 5
        s = s & "��"
        i = i + 1
    Loop
    MsgBox s
    
End Sub

Sub q58()
    Dim i As Long: i = 1
    Dim s As String
    Do While i <= 5
        s = s & CStr(i)
        i = i + 1
    Loop
    MsgBox s
    
End Sub

Sub q59()
    Dim i As Long: i = 5
    Dim s As String
    Do While i >= 1
        s = s & CStr(i)
        i = i - 1
    Loop
    MsgBox s
  
End Sub
Sub q60()
    Dim i As Long: i = 1
    Dim s As String
    Do While i <= 5
        s = s & i & "^2=" & (i * i) & vbCrLf
        i = i + 1
    Loop
    MsgBox s
    
End Sub


Sub q61()
    Dim i As Long: i = 1
    Dim line As String, s As String
    Do While i <= 5
        line = line & "��"
        s = s & line & "��"
        i = i + 1
    Loop
    MsgBox s
    
End Sub

Sub q62()
    Dim i As ling: i = 1
    Do While i <= 10
        Cells(i, "A").Value = "��"
            i = i + 1
    Loop
End Sub

Sub q63()
    Dim i As Long: i = 1
    Dim sum As Long
    Do While i <= 10
        If i Mod 2 = 1 Then sum = sum + i
        i = i + 1
    Loop
    MsgBox sum
    
'    �E�Ȃ�end if������Ȃ��H
'    �EIf .. Then�̂��Ƃɕϐ��̃C���N�������g�����邯�ǂȂ��H
    
End Sub

Sub q64()
    Dim i As Long: i = 1
    Do While i <= 10
        If Cells(i, 1).Value = "X" Then
            MsgBox "��������:" & Cells(i, 1).address(False, False)
            Exit Do
        End If
        i = i + 1
    Loop
    
End Sub

Sub q65()
    Dim i As Long: i = 1
    Do While i <= 10
        MsgBox i
        i = i + 1
        If i > 5 Then
            Exit Do
        End If
    Loop
    
End Sub

Sub q66()
    Dim i As Long
    i = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox i
    
End Sub


Sub q67()
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    MsgBox lastrow
    
End Sub

Sub q68()
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox Cells(lastrow, 1).Value
    
End Sub

Sub q69()
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
     Cells(lastrow + 1, 1).Value = "END"
    
End Sub

Sub q70()
    Dim i, lastrow As Long: i = 1
    Dim sum As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Do While i <= lastrow
        sum = sum + Cells(i, 1).Value
        i = i + 1
    Loop
    MsgBox sum
    
End Sub

Sub q71()
    Dim lastA, lastB, i As Long
    lastA = Cells(Rows.Count, 1).End(xlUp).Row
    lastB = Cells(Rows.Count, 2).End(xlUp).Row
    MsgBox "A��F" & Cells(lastA, 1).Value & vbCrLf & "B��F" & Cells(lastB, 2).Value
        
End Sub

Sub q72()
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 3).End(xlUp).Row
    MsgBox Cells(lastrow, 3).Value
    
End Sub

Sub q73()
    Dim i, lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastrow Step 1
        MsgBox Cells(i, 1).Value
    Next i
End Sub

Sub q74()
    Dim lastRows As Long
    lastrow = Cells(Row.Count, 1).End(xluo).Row
    Rows(lastrow).Delete
    
End Sub

Sub q75()
    Dim lastRows As Long, i As Long
    lastRows = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To 5
        Cells(lastRows + i, 1).Value = i
    Next i
    
End Sub

Sub q76()
    Dim price As Long
    Const tax As Double = 0.1
    price = 200
    MsgBox price * (1 + tax)

End Sub


Sub q77()
    Dim name As String, address As String
    name = "�R�c"
    address = "����"
    MsgBox name & ":" & address
    
End Sub

Sub q78()
    ' 10 + (20�~2) = 50
    Dim ans As Long
    ans = 10 + 20 * 2
    MsgBox ans
    
End Sub

Sub q79()
    Dim i As Long, total As Long
    For i = 1 To 5
        total = totak + 1
    Next i
    MsgBox total
    
End Sub

Sub q80()
    Dim i As Long, result As String
    For i = 1 To 5
        result = result & CStr(i)
    Next i
    MsgBox result
    
End Sub

Sub q81()
    Dim i As Long, total As Long: i = 1
    Do While i <= 5
        total = total + i
        i = i + 1
    Loop
    MsgBox total
    
End Sub


Sub q82()
    Dim i As Long, stars As String: i = 1
    Do Until i >= 5
        stars = stars & "��"
    Loop
    MsgBox stars
    
End Sub

Sub q83()
    Dim i As Long
    For i = 1 To 3
        MsgBox "For:" & i
    Next i
    
    i = 1
    Do While i <= 3
        MsgBox "Do:" & i
        i = i + 1
    Loop
    
End Sub


Sub q84()
    Dim i As Long: i = 1
    Do While i <= 10
        MsgBox i
        i = i + 1
        If i > 5 Then
            Exit Do
        End If
    Loop
    
End Sub

Sub q85()
    Dim a As Long, b As Long
    a = 1
    Do While a <= 3
        b = 1
        Do While b <= 3
          MsgBox a & "�~" & b & "=" & (a * b)
          b = b + 1
        Loop
    a = a + 1
    Loop
        
End Sub


Sub q86()
    Dim lastrow As Long, i As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastrow
        MsgBox Cells(i, 1).Value
    Next i
    
End Sub

Sub q87()
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lastrow + 1, 1).Value = "END"
    
End Sub

Sub q88()
    Dim lastrow As Long, i As Long, total As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastrow
        total = total + Cells(i, 1).Value
    Next i
    Cells(lastrow + 1, 1).Value = "���v�F" & total
    
    
End Sub

Sub q89()
    Dim lastrow As Long, i As Long: i = 1
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Do While i <= lastrow
        MsgBox Cells(i, 1).Value
        i = i + 1
        
    Loop
End Sub

Sub q90()
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Rows(lastrow).Delete
    
End Sub

Sub q91()
    Dim total As Double, lastrow As Long, i As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastrow
        total = total + Cells(i, 1).Value
    Next i
    MsgBox "���v" & total & "�~" & vbCrLf & "����" & (total / lastrow) & "�~"
    
End Sub

Sub q92()
    Dim lastrow As Long, i As Long: i = 1
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Do While i <= lastrow
        If Cells(i, 1).Value = "�R�c" Then
            MsgBox "���������s�ԍ��F" & i
            Exit Do
        End If
        i = i + 1
    Loop
    
End Sub

Sub q93()
    Dim lastrow As Long, i As Long: i = 1
    lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    Do While i <= lastrow
        MsgBox Cells(i, 2).Value & "�l�ł�"
        i = i + 1
    Loop
    
End Sub


Sub q94()
    Dim lastrow As Long, i As Long, total As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastrow
        If Cells(i, 1).Value Mod 2 = 0 Then
            total = total + Cells(i, 1).Value
        End If
    Next i
    MsgBox "���v" & total
    
        
End Sub


Sub q95()
    Dim i As Long
    For i = 1 To 5
        Cells(i, 1).Value = String(i, "��")
    Next i
    
End Sub


Sub q96()
    Dim lastrow As Long, i As Long: i = 1
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Do While i <= lastrow
        If Cells(i, 1).Value <= 70 Then
            MsgBox Cells(i, 1).Value & "�_�F" & "�s���i"
        Else
            MsgBox Cells(i, 1).Value & "�_�F" & "���i"
        End If
        i = i + 1
    Loop
            
End Sub

Sub q97()
    Dim lastrow As Long, i As Long, max As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastrow
        If max < Cells(i, 1).Value Then
            max = Cells(i, 1).Value
        End If
    Next i
    MsgBox "�ő�l�F" & max
    
End Sub

Sub q98()
    Dim lastrow As Long, i As Long, min As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    min = Cells(1, 1).Value
    For i = 1 To lastrow
        If min > Cells(i, 1).Value Then
            min = Cells(i, 1).Value
        End If
    Next i
        
    MsgBox "�ŏ��l�F" & min
    
End Sub

Sub q99()
    Dim lastrow As Long, i As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastrow
        If Cells(i, 1).Value >= 80 Then
            MsgBox Cells(i, 1).Value & "�_�F" & "�D"
        ElseIf Cells(i, 1).Value < 59 Then
            MsgBox Cells(i, 1).Value & "�_�F" & "�s��"
        Else
            MsgBox Cells(i, 1).Value & "�_�F" & "��"
        End If
    Next i
    
End Sub

Sub q100()
    Dim lastrow As Long, i As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastrow
        If Cells(i, 1) Mod 2 = 0 Then
            MsgBox Cells(i, 1).Value & "�F" & "����"
        Else
            MsgBox Cells(i, 1).Value & "�F" & "�"
        End If
    Next i
    
    
End Sub


'����


' vbCrLf�� �u���̒n�_������s�v

' String(��, "�o��") ��  �u�y���z�̕������y�o�́z���o�͂���v

' �Z���̔Ԓn���o�͂������ꍇ
'   cells(1, 1).Address(False, False) �o�́��uA1�v

' �f�[�^�̓������ŏI��̒l���擾
'   cells(lastRow, 1).Value

' �擾�����f�[�^����ŏI��̍s�ԍ����擾
'   �ϐ� = Cells(Rows.Count, 1).End(xiUp).Row



Sub try()
'    1.Long�^�̏����l�m�F
'    Dim s As Long
'    MsgBox s
    
'    2.s��Long�^�ɂ���CStr���Ȃ����Ă��Ȃ������ʂ��u0�v
'    Dim i, s As Long
'    For i = 5 To 1 - 1
'        s = s + i
'    Next i
'    MsgBox s
    

    
End Sub
