Attribute VB_Name = "Heard"
Option Explicit
Public Load As New Load
Public Mons As New Motion
Public Map As New Map
Public FerstLoad As New FerstLoad
Public X(15) As Single            ' ������� x ����. ������� ���������(�������)
Public Y(15) As Single            ' ������� y ����. ������� ���������(�������)
Public intPos(15) As Integer      ' ������� ������� ��������
Public intMaxPos(15) As Integer   ' ��� ������� ��������
Public strDir(15) As String       ' ����������� ��������
Public blnMove(15) As Boolean     ' ���������� �������� � ������� ������
Public Personag As Integer        ' �������� ��������
Public Catch As Boolean           '����, ������� ����, ��� ��������� �������
Public ServiseForm As Object      ' ������ �� ������� �����
Public objReturn As Object        ' ������ �� ��������� �����
Public SeeDistance(15) As Integer ' ������ ������
Public Xp As Integer, Yp As Integer  ' ����������� ���������
Public Const SetMajor = 11           ' ���������� �������� �� ���������
Public Place(15, 42, 3) As Byte   ' ���������� �� �������� (0-���, 1-��������� X ����������, 2-��������� Y ����������,3-���������� �� ������)
Public FurnW(223) As Byte        ' ������ ���������
Public FurnH(223) As Byte        ' ������ ���������
Public maxFurn As Integer            ' ��� ������� ������
Public GameDir As String               ' ����� ����
Public intCount As Integer

Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal X As Long, ByVal � As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long

Sub Main()

End Sub

