Attribute VB_Name = "MainOS"
Option Explicit
Public Const NumPic = 1
Public Const SetMajor = 11
Public picMon(NumPic, SetMajor, 1) As New StdPicture
Dim Mons As New Motion
Public intCount As Integer
Public Place(21, 62, 2) As Integer   ' ���������� �� �������� (0-���, 1-��������� X ����������
                                     '  2-��������� Y ����������)
Public FurnW(31) As Integer          ' ������ ���������
Public FurnH(31) As Integer          ' ������ ���������

Sub Main()
    OffSp.Show
End Sub

