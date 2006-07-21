Attribute VB_Name = "MainOS"
Option Explicit
Public Const NumPic = 1
Public Const SetMajor = 11
Public picMon(NumPic, SetMajor, 1) As New StdPicture
Dim Mons As New Motion
Public intCount As Integer
Public Place(21, 62, 2) As Integer   ' информация об объектах (0-вид, 1-начальная X координата
                                     '  2-начальная Y координата)
Public FurnW(31) As Integer          ' ширина интерьера
Public FurnH(31) As Integer          ' длинна интерьера

Sub Main()
    OffSp.Show
End Sub

