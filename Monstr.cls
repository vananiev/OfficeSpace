VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Motion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim imgX As ListImage
Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal X As Long, ByVal у As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long
Dim lngW As Long, lngH As Long
Dim lnX As Long, lnY As Long
Dim lngRtn As Long

Sub AddMons(Personag As Integer, Name As String)
DoEvents
Dim intCount As Integer
    On Error GoTo ExitSub
    frmMon.Mon.Picture = LoadPicture(GameDir & "\Image\Personag\" & Name & "_" & Personag & "_0.gif")
For intCount = 0 To SetMajor
        'загрузка картиноки и сохранение в imlPon
    frmMon.Mon.Picture = LoadPicture(GameDir & "\Image\Personag\" & Name & "_" & Personag & "_" & intCount & ".gif")
    frmMon.MonNeg.Picture = LoadPicture(GameDir & "\Image\Personag\NPersonag\" & Name & "_" & Personag & "_" & intCount & ".gif")
    'доб в набор позитев
    Set imgX = frmMon.imlMon.ListImages.Add(, , frmMon.Mon.Picture)
    'доб в набор  негатив
    Set imgX = frmMon.imlMon.ListImages.Add(, , frmMon.MonNeg.Picture)
        'загрузка картиноки и сохранение
    frmMon.Mon.Picture = LoadPicture(GameDir & "\Image\Personag\MPersonag\" & Name & "_" & Personag & "_" & intCount & ".gif")
    frmMon.MonNeg.Picture = LoadPicture(GameDir & "\Image\Personag\NMPersonag\" & Name & "_" & Personag & "_" & intCount & ".gif")
    'доб в набор позитев
    Set imgX = frmMon.imlMon.ListImages.Add(, , frmMon.Mon.Picture)
    'доб в набор  негатив
    Set imgX = frmMon.imlMon.ListImages.Add(, , frmMon.MonNeg.Picture)
Next intCount
ExitSub:
lnX = frmMon.imlMon.ListImages.Count
End Sub

Private Sub Move(X As Single, Y As Single, Direction As String)
Dim XCorection As Integer
Dim Ycorection As Integer
 ' определение направления движения
If Direction = "up" Then XCorection = 1: Ycorection = -1
If Direction = "doun" Then XCorection = -1: Ycorection = 1
If Direction = "left" Then XCorection = -1: Ycorection = -1
If Direction = "right" Then XCorection = 1: Ycorection = 1
If Direction = "stand" Then XCorection = 0: Ycorection = 0
 ' перемещаем  в новый участок
X = X + 4 * XCorection
Y = Y + 2 * Ycorection
End Sub

Public Sub Paint(Personag As Integer, X As Single, Y As Single, Position As Integer)
Dim intM As Integer, intN As Integer, mXp As Integer, mYp As Integer
Dim furX As Single, furY As Single
Dim Xstr As Integer, Ystr As Integer
Dim pPersonag As Integer
Map.DectoPlc X + 25, Y + 55, intN, intM
    ' цикл анимации
    frmMon.MonNeg.Picture = frmMon.imlMon.ListImages.Item(Position + 1 + 48 * Personag).Picture
    frmMon.Mon.Picture = frmMon.imlMon.ListImages.Item(Position + 48 * Personag).Picture
    For mXp = intN - 2 To intN + 1
        For mYp = intM + 1 To intM + 7
            If mYp > 42 Or mXp > 15 Or mXp < 0 Then GoTo NextSearch
            pPersonag = Place(mXp, mYp, 0)
            If pPersonag > 15 And pPersonag < 255 Then      ' поиск мебели
                Xstr = Place(mXp, mYp, 1)
                Ystr = Place(mXp, mYp, 2)
                If Xstr = mXp And Ystr = mYp Then
                    Map.PlstoDec mXp, mYp, furX, furY
                    furX = furX - 47 - X
                    furY = furY - 95 - Y
                    frmMon.FurnNeg.Picture = frmMon.imlFurn.ListImages.Item((pPersonag - 16) * 2 + 2).Picture
                    frmMon.Furn.Picture = frmMon.imlFurn.ListImages.Item((pPersonag - 16) * 2 + 1).Picture
                    ' применяем маску для мебели
                    lngRtn = BitBlt(frmMon.Mon.hDC, furX, furY, 150, 185, frmMon.FurnNeg.hDC, 0, 0, vbSrcAnd)
                    ' рисуем
                    lngRtn = BitBlt(frmMon.Mon.hDC, furX, furY, 150, 185, frmMon.Furn.hDC, 0, 0, vbSrcPaint)
                End If
            End If
NextSearch:
            Next mYp
        Next mXp
    ' применяем маску
    lngRtn = BitBlt(ServiseForm.hDC, X, Y, 50, 75, frmMon.MonNeg.hDC, 0, 0, vbSrcAnd)
    ' рисуем
    lngRtn = BitBlt(ServiseForm.hDC, X, Y, 50, 75, frmMon.Mon.hDC, 0, 0, vbSrcPaint)
End Sub

Public Sub Fps()
Dim lngRtn As Long
    ' прорисовка плитки и мебели
    ServiseForm.Picture = ServiseForm.imlFon.ListImages.Item(1).Picture
    For Xp = 0 To 15
        For Yp = 6 To 42
            Personag = Place(Xp, Yp, 0)
            If Personag < 16 Then  ' поиск персонажей
                Paint Personag, X(Personag), Y(Personag), intPos(Personag)
            End If
            If Personag >= 27 And Personag <= 31 Then ' поиск бонусов
                Map.FurnPaint Xp, Yp, (Personag - 16) * 2 + 1
            End If
        Next Yp
    Next Xp
End Sub

Public Function PersonagMove(Personag As Integer) As Byte
If intMaxPos(Personag) - intPos(Personag) = 21 Then
    If Not (CheckingMove(Personag, X(Personag), Y(Personag), PersonagMove)) Then
        blnMove(Personag) = False
        strDir(Personag) = "stand"
        Exit Function
    Else
        blnMove(Personag) = True
    End If
End If
Move X(Personag), Y(Personag), strDir(Personag)
intPos(Personag) = intPos(Personag) + 4
If intPos(Personag) >= intMaxPos(Personag) Then
    blnMove(Personag) = False
    intPos(Personag) = intPos(Personag) - 4
    strDir(Personag) = "stand"
End If
End Function

Public Sub MonstrsMove(NumPersonag As Integer)
Dim MMPersonag As Integer
For MMPersonag = 1 To NumPersonag
    If strDir(MMPersonag) = "stand" Then MonsMove MMPersonag, X(MMPersonag), Y(MMPersonag)
    If intMaxPos(MMPersonag) - intPos(MMPersonag) = 21 Then
        If Not (CheckingMove(MMPersonag, X(MMPersonag), Y(MMPersonag), 0)) Then
            blnMove(MMPersonag) = False
            strDir(MMPersonag) = "stand"
            GoTo EndPnt
        Else
            blnMove(MMPersonag) = True
        End If
    End If
    Move X(MMPersonag), Y(MMPersonag), strDir(MMPersonag)
    intPos(MMPersonag) = intPos(MMPersonag) + 4
    If intPos(MMPersonag) >= intMaxPos(MMPersonag) Then
        blnMove(MMPersonag) = False
        intPos(MMPersonag) = intPos(MMPersonag) - 4
        strDir(MMPersonag) = "stand"
    End If
EndPnt:
Next MMPersonag
End Sub

Private Sub MonsMove(Personag As Integer, xs As Single, ys As Single)
Dim N As Integer
Dim Xstp As Integer, Ystp As Integer
Dim NumFurn As Integer
Catch = False
Map.DectoPlc xs + 25, ys + 55, Xstp, Ystp
'Охота
For Xp = Xstp - SeeDistance(Personag) To Xstp + SeeDistance(Personag)
    For Yp = Ystp - Personag To Ystp + Personag
        If Xp < 0 Or Yp < 0 Or Xp > 15 Or Yp > 48 Then GoTo NextPos
        NumFurn = Place(Xp, Yp, 0)
     '  поиск 0 персонажа
        If NumFurn = 0 Then
            If Abs(Xstp - Xp) <= 1 And Abs(Ystp - Yp) <= 1 Then Catch = True
            If Yp - Ystp < 0 And Xp - Xstp < 0 Then N = 1: GoTo Napr
            If Yp - Ystp < 0 And Xp - Xstp > 0 Then N = 3: GoTo Napr
            If Yp - Ystp > 0 And Xp - Xstp < 0 Then N = 2: GoTo Napr
            If Yp - Ystp > 0 And Xp - Xstp > 0 Then N = 0: GoTo Napr
            GoTo Napr
        End If
NextPos:
    Next Yp
Next Xp

'Беспорядочное движение
    Randomize Timer
    N = Rnd * (20 + Personag \ 3)
Napr:
    If N = 0 Then strDir(Personag) = "right": intPos(Personag) = 1: intMaxPos(Personag) = 22
    If N = 1 Then strDir(Personag) = "left": intPos(Personag) = 27: intMaxPos(Personag) = 48
    If N = 2 Then strDir(Personag) = "doun": intPos(Personag) = 3: intMaxPos(Personag) = 24
    If N = 3 Then strDir(Personag) = "up": intPos(Personag) = 25: intMaxPos(Personag) = 46
    If N = 4 Then strDir(Personag) = "stand": intPos(Personag) = intPos(Personag) - 4: intMaxPos(Personag) = 22
End Sub

Private Function CheckingMove(Personag As Integer, X As Single, Y As Single, Bonus As Byte) As Boolean
Dim XCorection As Integer
Dim Ycorection As Integer
' убирание места положения персонажа
    Map.DectoPlc X + 25, Y + 55, Xp, Yp
    Place(Xp, Yp, 0) = 255
' определение направления движения
If strDir(Personag) = "up" Then XCorection = 1: Ycorection = -1
If strDir(Personag) = "doun" Then XCorection = -1: Ycorection = 1
If strDir(Personag) = "left" Then XCorection = -1: Ycorection = -1
If strDir(Personag) = "right" Then XCorection = 1: Ycorection = 1
If strDir(Personag) = "stand" Then XCorection = 0: Ycorection = 0
' перевод из декартовой сис.
Map.DectoPlc X + 27 + 24 * XCorection, Y + 55 + 12 * Ycorection, Xp, Yp
If Place(Xp, Yp, 0) < 16 Or Place(Xp, Yp, 0) > 31 And Place(Xp, Yp, 0) < 255 Or Place(Xp, Yp, 3) = 0 Or Personag <> 0 And Place(Xp, Yp, 0) > 26 And Place(Xp, Yp, 0) < 255 Then
ExitSub:
    CheckingMove = False
    ' установка положения персонажа
    Map.DectoPlc X + 27, Y + 55, Xp, Yp
    Place(Xp, Yp, 0) = Personag
    Bonus = 0
Else
    CheckingMove = True
    ' установка положения персонажа
    Bonus = Place(Xp, Yp, 0)
    Place(Xp, Yp, 0) = Personag
End If
End Function

 Sub KeyPress(Personag As Integer, KeyCode As Integer, Shift As Integer)
    If blnMove(Personag) Then Exit Sub
    If KeyCode = 83 Then
        strDir(Personag) = "right"
        intPos(Personag) = 1
        intMaxPos(Personag) = 22
        blnMove(Personag) = True
    End If
    If KeyCode = 65 Then
        strDir(Personag) = "doun"
        intPos(Personag) = 3
        intMaxPos(Personag) = 24
        blnMove(Personag) = True
    End If
    If KeyCode = 87 Then
        strDir(Personag) = "up"
        intPos(Personag) = 25
        intMaxPos(Personag) = 46
        blnMove(Personag) = True
    End If
    If KeyCode = 81 Then
        strDir(Personag) = "left"
        intPos(Personag) = 27
        intMaxPos(Personag) = 48
        blnMove(Personag) = True
    End If
End Sub

