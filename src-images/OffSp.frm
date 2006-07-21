VERSION 5.00
Begin VB.Form OffSp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   21
      Left            =   0
      Top             =   3720
   End
End
Attribute VB_Name = "OffSp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mons As New OfficeSpace.Motion
Dim Map As New Map
Dim x(7) As Single            ' текущая x корд. данного персонажа(пикселы)
Dim y(7) As Single            ' текущая y корд. данного персонажа(пикселы)
Dim intPos(7) As Integer      ' текущая позиция картинки
Dim intMaxPos(7) As Integer   ' мах позиция картинки
Dim strDir(7) As String       ' направление движения
Dim blnMove(7) As Boolean     ' выполнение движения в текущий момент
Dim Personag As Integer    ' активный персонаж
Dim Xp As Single, Yp As Single
Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal x As Long, ByVal у As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long
Dim lngRtn As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not (blnMove(0)) Then
If KeyCode = 83 Then
    tmrMove.Enabled = True
    strDir(0) = "right"
    intPos(0) = 1
    intMaxPos(0) = 22
    blnMove(0) = True
End If
If KeyCode = 65 Then
    tmrMove.Enabled = True
    strDir(0) = "doun"
    intPos(0) = 3
    intMaxPos(0) = 24
    blnMove(0) = True
End If
If KeyCode = 87 Then
    tmrMove.Enabled = True
    strDir(0) = "up"
    intPos(0) = 25
    intMaxPos(0) = 46
    blnMove(0) = True
End If
If KeyCode = 81 Then
    tmrMove.Enabled = True
    strDir(0) = "left"
    intPos(0) = 27
    intMaxPos(0) = 48
    blnMove(0) = True
End If
End If
End Sub

Private Sub Form_Load()
    Dim n As Integer
    Dim m As Integer
    For n = 0 To Screen.Height Step 24
        Line (0, n)-(1200, n - 600)
    Next n
    For n = -600 To Screen.Width Step 24
        Line (0, n)-(1200, n + 600)
    Next n
    ' начальное состояние оффиса
    For n = 0 To 21
        For m = 0 To 62
            Place(n, m, 0) = 16
        Next m
    Next n
    LoadPersonag
    LoadFurniture
    frmMon.Show
End Sub

Private Sub tmrMove_Timer()
For Personag = 0 To 1
         If Personag <> 0 And Not (blnMove(Personag)) Then MonstrsMove Personag, x(Personag), y(Personag): blnMove(Personag) = True
        If intMaxPos(Personag) - intPos(Personag) = 21 Then
            If Not (CheckingMove(x(Personag), y(Personag))) Then blnMove(Personag) = False: Exit Sub
        End If
            Mons.Move Personag, x(Personag), y(Personag), intPos(Personag), strDir(Personag)
            CheckingPaint x(Personag), y(Personag)
            Refresh
            intPos(Personag) = intPos(Personag) + 4
            If intPos(Personag) >= intMaxPos(Personag) Then blnMove(Personag) = False: intPos(Personag) = intPos(Personag) - 4: strDir(Personag) = "stand"
Next Personag
End Sub

Private Function CheckingMove(x As Single, y As Single) As Boolean
Dim XCorection As Integer
Dim YCorection As Integer
' определение направления движения
If strDir(0) = "up" Then XCorection = 1: YCorection = -1
If strDir(0) = "doun" Then XCorection = -1: YCorection = 1
If strDir(0) = "left" Then XCorection = -1: YCorection = -1
If strDir(0) = "right" Then XCorection = 1: YCorection = 1
' перевод из декартовой сис.
Map.DectoPlc x + 5 + 24 * XCorection, y + 62 + 12 * YCorection, Xp, Yp
If Place(Xp, Yp, 0) > 31 Then
    CheckingMove = False
Else
    CheckingMove = True
End If
End Function

Public Sub LoadPersonag()
    'координаты персонажей
    x(0) = 47
    y(0) = 35
    x(1) = 677
    y(1) = 299
    'сохраняем фон под персонажами и прорисовываем их в 1 раз
    For Personag = 0 To 1
        blnMove(Personag) = False
        Map.DectoPlc x(Personag) + 5, y(Personag) + 62, Xp, Yp
        Place(Xp, Yp, 0) = 0                           '=бос
        lngRtn = BitBlt(frmMon.picBack(Personag + 1).hDC, 0, 0, 50, 75, OffSp.hDC, x(Personag), y(Personag), vbSrcCopy)
        Mons.Move Personag, x(Personag), y(Personag), 9, "right"
    Next Personag
End Sub

Public Sub LoadFurniture()
    Dim X0 As Single
    Dim Y0 As Single
    Dim NumFurn As Integer  ' номер интерьера (0<=x<=31)
    'загрузка интерьера
    PlaceMassiv 10, 10, 16
    PlaceMassiv 7, 20, 17
    PlaceMassiv 5, 30, 20
End Sub

Private Sub CheckingPaint(xs As Single, ys As Single)
Dim Xstp As Single
Dim Ystp As Single
Dim Xp As Single
Dim Yp As Single
Dim NumFurn As Integer
Dim m As Integer
Dim n As Integer
For m = 1 To 7
    For n = -1 To 1 Step 2
        Map.DectoPlc xs + n * 24 + 10, ys + m * 12 - 2 + 62, Xstp, Ystp
        NumFurn = Place(Xstp, Ystp, 0)
        If NumFurn > 23 Then             ' поиск мебели
            Xp = Place(Xstp, Ystp, 1)
            Yp = Place(Xstp, Ystp, 2)
            Map.FurnPaint Int(Xp), Int(Yp), (NumFurn - 16) * 2 + 1
        End If
        If NumFurn < 16 Then  ' поиск персонажей
            Mons.Move NumFurn, x(NumFurn), y(NumFurn), intPos(NumFurn), strDir(NumFurn)
        End If
    Next n
Next m
End Sub

Public Sub PlaceMassiv(X0 As Single, Y0 As Single, NumFurn As Integer)
    Dim Xstp As Single
    Dim Ystp As Single
    Dim intWork As Integer
    Dim n As Integer
    Xstp = X0
    Ystp = Y0
    Map.PlstoDec Xstp, Ystp, Xp, Yp
         ' up side
    For n = 1 To FurnW(NumFurn) - 1
        Map.DectoPlc Xp + 1, Yp + 2, Xstp, Ystp
        Place(Xstp, Ystp, 0) = NumFurn + 16 ' =интерьер
        Place(Xstp, Ystp, 1) = Int(X0)
        Place(Xstp, Ystp, 2) = Int(Y0)
        Yp = Yp - 12
        Xp = Xp + 24
    Next n
        'right side
    For n = 1 To FurnH(NumFurn) - 1
        Map.DectoPlc Xp + 1, Yp + 2, Xstp, Ystp
        Place(Xstp, Ystp, 0) = NumFurn + 16 ' =интерьер
        Place(Xstp, Ystp, 1) = Int(X0)
        Place(Xstp, Ystp, 2) = Int(Y0)
        Yp = Yp + 12
        Xp = Xp + 24
    Next n
        'doun side
    For n = 1 To FurnW(NumFurn) - 1
        Map.DectoPlc Xp + 1, Yp + 2, Xstp, Ystp
        Place(Xstp, Ystp, 0) = NumFurn + 16 ' =интерьер
        Place(Xstp, Ystp, 1) = Int(X0)
        Place(Xstp, Ystp, 2) = Int(Y0)
        Yp = Yp + 12
        Xp = Xp - 24
    Next n
        'left side
    For n = 1 To FurnH(NumFurn) - 1
        Map.DectoPlc Xp + 1, Yp + 2, Xstp, Ystp
        Place(Xstp, Ystp, 0) = NumFurn + 16 ' =интерьер
        Place(Xstp, Ystp, 1) = Int(X0)
        Place(Xstp, Ystp, 2) = Int(Y0)
        Yp = Yp - 12
        Xp = Xp - 24
    Next n
    ' прорисовка интерьера
    Map.FurnPaint Int(X0), Int(Y0), (NumFurn * 2) + 1

End Sub

Private Sub MonstrsMove(Personag As Integer, xs As Single, ys As Single)
Dim n As Integer
Dim m As Integer
Dim Xstp As Single
Dim Ystp As Single
Dim NumFurn As Integer
'Охота
For m = (-Personag - 2) To (Personag + 2)
    For n = (-Personag - 4) To (Personag + 4) Step 2
        Map.DectoPlc xs + n * 24 + 10, ys + m * 12 - 2 + 62, Xstp, Ystp
        NumFurn = Place(Xstp, Ystp, 0)
     '  поиск 0 персонажа
        If NumFurn = 0 Then
            If m < 0 And n < 0 Then n = 1: GoTo Napr
            If m < 0 And n > 0 Then n = 3: GoTo Napr
            If m > 0 And n < 0 Then n = 2: GoTo Napr
            If m > 0 And n > 0 Then n = 0: GoTo Napr
        End If
    Next n
Next m

'Беспорядочное движение
    Randomize Timer
    n = Rnd * 5
Napr:
    If n = 0 Then strDir(Personag) = "right": intPos(Personag) = 1: intMaxPos(Personag) = 22
    If n = 1 Then strDir(Personag) = "left": intPos(Personag) = 27: intMaxPos(Personag) = 48
    If n = 2 Then strDir(Personag) = "doun": intPos(Personag) = 3: intMaxPos(Personag) = 24
    If n = 3 Then strDir(Personag) = "up": intPos(Personag) = 25: intMaxPos(Personag) = 46
    If n > 3 Then strDir(Personag) = "stand": intPos(Personag) = intPos(Personag) - 4: intMaxPos(Personag) = 22
End Sub
