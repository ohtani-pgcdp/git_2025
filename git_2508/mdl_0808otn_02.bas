Attribute VB_Name = "mdl_0808otn_02"
Option Explicit


'*********************************************************
'* ForNextの練習
'*********************************************************

Sub mcr0808_01()
  Dim i               '変数iを宣言
  For i = 1 To 10    '変数iを１から10まで１ずつカウントアップ
    Cells(i, 1) = 1  'A列のi行目のセルに1を入れる
  Next i               'Forの範囲はここまで


End Sub

Sub mcr0808_02()
  Dim i
  For i = 10 To 1 Step -1   '変数iを１から10まで１ずつカウントアップ
    Cells(i, 3) = 1  'A列のi行目のセルに1を入れる
  Next i
End Sub

Sub mcr0808_03()
 Dim i
  For i = 10 To 1 Step -1   '変数iを１から10まで１ずつカウントアップ
    Cells(i, 7) = 1  'A列のi行目のセルに1を入れる
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
    
'「<>」は、「○○ではない」ということを調べる演算子です。
'「<>」の左と右を比べて、等しくないことを調べる比較演算子です。



End Sub

' 調べる：インクリメントはどこで行われてる？

Sub mcr0808_06()
  Dim i, j
      For i = 15 To 25
          For j = 1 To 10
              Cells(i, j) = 1
          Next j
      Next i

End Sub

'１問目
'以下を３回、メッセージボックスに表示させる。
'「５５」を３回表示

Sub mcr0808_07()
  Dim i
    For i = 1 To 3
      MsgBox 55
    Next i
    
End Sub

'２問目
'以下をメッセージボックスに表示
'※なお､①とは1回目という意味
'①5
'②10
'③15

Sub mcr0808_08()
  Dim i
    For i = 5 To 15 Step 5
      MsgBox i
    Next i
    
End Sub


