Attribute VB_Name = "MainOS"
Option Explicit
Public Mons As New GMotion.Motion
Public Map As New GMotion.Map
Public Load As New GMotion.Load
Public Member As New GMotion.Member
Public SEvents As New SEvents
Public intFps As Integer          ' Fps
Public Company As String          '��� ��������
Public Level As Integer           '����� ������
Public NumPersonag As Integer     '���. ����������
Public Files As Byte           '���. �����
Public Xp As Integer, Yp As Integer  ' ����������� ���������
Public Lifes As Byte                  '�����
Public Money As Integer              '���. �����
Public startLifes As Byte                 '�� ������ ������ �����
Public startMoney As Integer              '�� ������ ������ ���. �����
Public SaveMode As Boolean           '��� �������� ��������

Sub Main()
    Member.MCatch = False
    intFps = 24
    frmMain.Show
End Sub
