Attribute VB_Name = "Module1"
Option Explicit

Public Sub fornex1()
'-----------------------------------------------------------------------------
'   数字の分だけ文字列を表示させる
'   Date     : 2021/01/19
'   Author   : N.Hoshii
'   URL      : http://psa2.kuciv.kyoto-u.ac.jp/staff/susaki/c/for.html
'   Modified : 新規作成(2021/01/19)
'-----------------------------------------------------------------------------
    Dim i As Integer    'カウンタ用変数
    Dim j As Integer    '数値入力ボックス用変数
    j = InputBox("数字を1つ入力して下さい")
    
    For i = 1 To 5
        MsgBox "" & j & """回目の表示"""
    Next i

End Sub
