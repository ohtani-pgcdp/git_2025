Attribute VB_Name = "test_0801_otn"
Option Explicit

Sub mcr01()
'  A1セルに「こんにちは」と表示させるマクロを作成しなさい。
  Cells(1, 1).Value = "こんにちは"
End Sub


Sub mcr03()
'  数値型の変数priceを宣言して「1200」を代入し、A5セルに出力させるマクロを作成しなさい。
  Dim price As Long
  price = 1200
  Cells(5, 1).Value = price
End Sub

Sub mcr05()
  '数値「2000」を定数「Price」として宣言し、A2セルに表示するマクロを作成しなさい。
  Const price As Long = 2000
  Cells(2, 1).Value = price
  
End Sub


Sub mcr08()
 '変数scoreに85を代入し、Cellsを使ってB5セルに出力するマクロを作成しなさい。
 Dim score As Long
 score = 85
 Cells(5, 2) = score
 
End Sub


Sub mcr10()
 '定数「BASIC」を80、変数「Add」を20として、合計点を計算しC1セルに表示するマクロを作成しなさい。
 
End Sub
