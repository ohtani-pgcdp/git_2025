Attribute VB_Name = "HandsOn"
Option Explicit

Sub RightUp_01()
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 1 To i
            Cells(i, j).Value = "■"
        Next j
    Next i
    
'cells(1,1)は■1つ、cells(2,2)は2つ...と行(列)番号と同じ数入力
'A列から順に増える挙動にする(1→2→3...)ためjの条件式右辺はi
           
End Sub

Sub LeftUp_02()
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 11 - i To 10
            Cells(i, j).Value = "■"
        Next j
    Next i
    
'・iが行、jが列を担当
'・iとjが11になるよう操作  cells(1, 10)→cells(2, 9)

'・「jが10になるまで」という条件を満たしつつ、■を入力するセルを指定するため
'この条件を満たすには
'1.行/列番号の最大値(今回は10)より大きな数字をForの条件に


      
End Sub

Sub RightDown_03()
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 1 To 11 - i
            Cells(i, j).Value = "■"
        Next j
    Next i
    
End Sub


Sub LeftDown_04()
  Dim i As Long, j As Long
  For i = 1 To 10
      For j = i To 10
          Cells(i, j).Value = "■"
      Next j
  Next i
  
  
'1行目はすべて■、そこから１つずつ減らすのでi = 1 to 10
'iが１ずつ増えていくので、■の入力を開始したい列の列番号としてjの条件式はiからスタート(j = i to 10)
  
End Sub

Sub RightHalf_05()
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 1 To i
            Cells(i, j).Value = "■"
        Next j
    Next i
    
    For i = 11 To 19
        For j = 1 To 20 - i
            Cells(i, j).Value = "■"
        Next j
    Next i
    
'10行目まではRightUp_01と同じ
'11行目がA列から■9個、そこから19行目まで1ずつ減るので外側はi = 11 to 19
'cells(11, 1)~cells(11, 9)まで入力したい&iを使って行数も制御する→jは1からスタート
'条件式の右辺を9にすると四角形に■が出力されてしまうので、空白としたいセル(J11,I12...)をループを重ねる毎に増やすため外側のiを右辺とする


   
End Sub


Sub LeftHalf_06()
    Dim i As Long, j As Long
    'ネスト1
    For i = 1 To 10
        For j = 11 - i To 10
            Cells(i, j).Value = "■"
        Next j
    Next i

    'ネスト2
    For i = 11 To 19
        For j = i - 9 To 10
            Cells(i, j).Value = "■"
        Next j
    Next i

    
'10行目まではLeftUp_02と同じ

'ネスト2の行を制御するiは、ネスト1までと対称とするため11 to 19
'ネスト2はB11(cells(11, 2))から■を入力したいため2を左辺、10を右辺としたい
'行を制御して入力しない列(11行目はA、12行目はAB...)を作るため、左辺が1ずつ減るようi(11 to 19) - 9とした

    
End Sub


Sub UpHalf_07()
'    Dim i As Long, j As Long
'    For i = 1 To 10
'        For j = 11 - i To 10
'            Cells(i, j).Value = "■"
'        Next j
'    Next i
'
'    For i = 1 To 10
'        For j = 10 + i To 19
'            Cells(i, j).Value = "■"
'        Next j
'    Next i


    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 11 - i To 10
            Cells(i, j).Value = "■"
        Next j
    Next i
    
    For i = 2 To 10
        For j = 11 To 9 + i
           Cells(i, j).Value = "■"
        Next j
    Next i
    
End Sub


Sub DouwHalf_08()
    Dim i As Long, j As Long
    
    For i = 1 To 10
        For j = 0 + i To 10
             Cells(i, j).Value = "■"
        Next j
    Next i
'jを0からスタートすることでi(列番号)と同じスタート位置に
    
    For i = 1 To 10
        For j = 11 To 20 - i
            Cells(i, j).Value = "■"
        Next j
    Next i
'K列(Cells(1, 11))から上限を１ずつ減らしたい
'ので、左辺を11、右辺を20 - iとして変化させる
    
End Sub


Sub AALDiamond_09()
'左上
    Dim i As Long, j As Long
    For i = 1 To 10
        For j = 11 - i To 10
            Cells(i, j).Value = "■"
        Next j
    Next i
    
'右上
    For i = 2 To 10
        For j = 11 To 9 + i
            Cells(i, j).Value = "■"
        Next j
    Next i
    
'左下
    For i = 11 To 20
        For j = i - 9 To 10
            Cells(i, j).Value = "■"
        Next j
    Next i
    
'右下
    For i = 11 To 19
        For j = 11 To 29 - i
            Cells(i, j).Value = "■"
        Next j
    Next i
        
End Sub


Sub test()
    Dim i As Long, j As Long
'    For i = 1 To 10
'        Cells(i, 1) = "■"
'    Next i
    For j = 1 To 10
        Cells(1, j) = "■"
    Next j
    
End Sub
