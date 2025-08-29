Attribute VB_Name = "test_0801_otn"
Option Explicit

Sub mcr01()
'  A1セルに「こんにちは」と表示させるマクロを作成しなさい。
  Cells(1, 1).Value = "こんにちは"
End Sub

Sub mcr02()
'  Cellsを使って、B3セルに「VBAテスト」と表示させなさい。
  Cells(3, 2).Value = "VBAテスト"
End Sub


Sub mcr03()
'  数値型の変数priceを宣言して「1200」を代入し、A5セルに出力させるマクロを作成しなさい。
  Dim price As Long
  price = 1200
  Cells(5, 1).Value = price
End Sub

Sub mcr04()
'  A1セルに「氏名」、B1セルに「点数」と入力するマクロを作成しなさい。
  Cells(1, 1).Value = "氏名"
  Cells(2, 1).Value = "点数"
  
End Sub


Sub mcr05()
  '数値「2000」を定数「Price」として宣言し、A2セルに表示するマクロを作成しなさい。
  Const price As Long = 2000
  Cells(2, 1).Value = price
  
End Sub

Sub mcr06()
'  変数greetingに「おはよう」を代入し、B5セルに出力するマクロを作成しなさい。
  Dim greeting As String
  greeting = "おはよう"
  Cells(5, 2).Value = greeting
  
End Sub

Sub mcr07()
'  変数を経由してセルに値を表示する基本的なマクロを作成しなさい｡
  Dim today As String
  today = "8月1日"
  Cells(1, 3) = today
  
End Sub


Sub mcr08()
 '変数scoreに85を代入し、Cellsを使ってB5セルに出力するマクロを作成しなさい。
 Dim score As Long
 score = 85
 Cells(5, 2).Value = score
 
End Sub
Sub mcr09()
'  変数firstNameに「山田」、lastNameに「太郎」を代入し、それらを結合してA1セルに出力するマクロを作成しなさい。
  Dim firstName As String, lastName As String
  firstName = "山田"
  lastName = "太郎"
  Cells(1, 1).Value = firstName & lastName
  
End Sub

Sub mcr10()
 '定数「BASIC」を80、変数「Add」を20として、合計点を計算しC1セルに表示するマクロを作成しなさい。
 Const BASIC As Long = 80
 Dim add As Long
 add = 20
 Cells(1, 3).Value = BASIC + add
End Sub
