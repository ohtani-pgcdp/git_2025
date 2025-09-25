Attribute VB_Name = "AutoGrader"
Option Explicit

'単一ブック(60シート)の採点フォント整備済み
'   フォント、単一ブックに対する処理は問題なし。次は複数ブック対応へ

Sub ProcessScoreForSheet()
   Dim answer As Range, correct As Range, cC As Double, _
   i As Long, a As Long, rate As Double, wb As Workbook
    
    Workbooks("01.Excelの基本操作_全60問.xlsx").Activate
    
    
    '採点用シートを作成・挿入↓
    Sheets.Add Sheets(1)
    Sheets(1).Name = "点数シート"
    Cells(1, 1).Value = "問題番号"
    Cells(1, 2).Value = "解答"
    Cells(1, 3).Value = "正答"
    Cells(1, 4).Value = "判定"
    Range("A1:D1").Font.Bold = True
    Range("A1:D1").HorizontalAlignment = xlCenter
    Range("A1:D1").Interior.Color = vbBlack
    Range("A1:D1").Font.Color = vbWhite
    
    For a = 2 To 61
        Cells(a, 1).Value = "Q" & (a - 1)
        Cells(a, 1).Font.Bold = True
    Next a
    
    For i = 2 To 61
        '解答を保持
        Set answer = Workbooks("01.Excelの基本操作_全60問.xlsx").Sheets(i).Cells(20, 1) '生徒が提出したファイル
        
        '正答を保持
        Workbooks("解答集(最終課題用).xlsx").Sheets(1).Activate '解答が書かれたファイル
        Set correct = Sheets(1).Cells(i - 1, 2)
        
        '元のブックに戻って採点開始
        Workbooks("01.Excelの基本操作_全60問.xlsx").Activate
        Cells(i, 2).Value = answer
        Cells(i, 3).Value = correct
        
        If answer = correct Then
            Cells(i, 4).Value = "○"
            Cells(i, 4).Font.Bold = True
            Range(Cells(i, 1), Cells(i, 4)).Interior.Color = vbRed
            Range(Cells(i, 1), Cells(i, 4)).Font.Color = vbWhite
            Cells(i, 4).HorizontalAlignment = xlCenter
            cC = cC + 1
        ElseIf answer = "" Then
            Cells(i, 4).Value = "未回答"
            Cells(i, 4).HorizontalAlignment = xlCenter
        Else
            Cells(i, 4).Value = "×"
            Cells(i, 4).Font.Bold = True
            Range(Cells(i, 1), Cells(i, 4)).Interior.Color = vbBlue
            Range(Cells(i, 1), Cells(i, 4)).Font.Color = vbWhite
            Cells(i, 4).HorizontalAlignment = xlCenter
        End If
    Next i
    
    Cells(62, 3).Value = "正答率："
    rate = (cC / 60 * 100)
    If rate = 100 Then
        Cells(62, 4).Font.Bold = True
        Cells(62, 4).Value = "全問正解！"
        Cells(62, 4).Interior.Color = vbYellow
        Cells(62, 4).HorizontalAlignment = xlCenter
    Else
        Cells(62, 4).Font.Bold = True
        Cells(62, 4).Value = rate & "%"
        Cells(62, 4).Interior.Color = vbYellow
        Cells(62, 4).HorizontalAlignment = xlCenter
    End If
    
    '最後にまとめて罫線を書き足し表にする
    Cells(1, 4).Borders(xlEdgeRight).Weight = xlMedium
    Range("A1:D61").Borders(xlInsideHorizontal).Weight = xlMedium
    Range("A1:D61").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("D2:D61").Borders(xlEdgeRight).Weight = xlMedium
    Range("A61:D61").Borders(xlEdgeBottom).Weight = xlMedium
    Range("C62:D62").Borders(xlEdgeLeft).Weight = xlMedium
    Range("C62:D62").Borders(xlEdgeBottom).Weight = xlMedium
    Range("C62:D62").Borders(xlEdgeRight).Weight = xlMedium
    Cells(62, 3).Borders(xlEdgeRight).LineStyle = xlContinuous
    
    '保存して閉じる
'    Set wb = Workbooks("01.Excelの基本操作_全60問.xlsx")
'    wb.Close savechanges:=True
    
End Sub
