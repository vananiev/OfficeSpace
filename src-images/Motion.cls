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
Attribute VB_Exposed = False
Option Explicit
Dim imgX As ListImage
Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal x As Long, ByVal у As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long
Dim lngW As Long, lngH As Long
Dim x As Long, y As Long

Public Sub AddMon(NumPic As Byte, Name As String)
Dim Major As Integer
For intCount = 0 To SetMajor
    'загрузка картиноки и сохранение в imlPon
    frmMon.Mon = LoadPicture(App.Path & "\" & Replace(Name & "_" & Str(Major) & ".gif", " ", ""))
    frmMon.Mon.Refresh
    frmMon.MonNeg.Refresh
    lngW = frmMon.Mon.ScaleWidth
    lngH = frmMon.Mon.ScaleHeight
    For x = 0 To lngW
    For y = 0 To lngH
    'негатив
    If frmMon.Mon.Point(x, y) = 16777215 Then
        frmMon.MonNeg.PSet (x, y), 16777215
    Else
        frmMon.MonNeg.PSet (x, y), 0
    End If
    Next y
    Next x
    'позитив
    For x = 0 To lngW
    For y = 0 To lngH
        If frmMon.Mon.Point(x, y) = 16777215 Then frmMon.Mon.PSet (x, y), 0
    Next y
    Next x
    'доб в набор позитев
    frmMon.Mon.Picture = frmMon.Mon.Image
    Set imgX = frmMon.imlMon.ListImages.Add(, , frmMon.Mon.Picture)
    'доб в набор  негатив
    frmMon.MonNeg.Picture = frmMon.MonNeg.Image
    Set imgX = frmMon.imlMon.ListImages.Add(, , frmMon.MonNeg.Picture)
    DoMerrir
    Major = Major + 1
Next intCount
End Sub

Public Sub Move(Personag As Integer, x As Single, y As Single, Position As Integer, Direction As String)
Dim lngRtn As Long
Dim XCorection As Integer
Dim YCorection As Integer
' определение направления движения
If Direction = "up" Then XCorection = 1: YCorection = -1
If Direction = "doun" Then XCorection = -1: YCorection = 1
If Direction = "left" Then XCorection = -1: YCorection = -1
If Direction = "right" Then XCorection = 1: YCorection = 1
If Direction = "stand" Then XCorection = 0: YCorection = 0
' цикл анимации
frmMon.MonNeg.Picture = frmMon.imlMon.ListImages.Item(Position + 1 + 48 * Personag).Picture
frmMon.Mon.Picture = frmMon.imlMon.ListImages.Item(Position + 48 * Personag).Picture
' восстанавливаем фон
lngRtn = BitBlt(OffSp.hDC, x, y, 50, 75, frmMon.picBack(Personag + 1).hDC, 0, 0, vbSrcCopy)
' перемещаем НЛО в новый участок
x = x + 4 * XCorection
y = y + 2 * YCorection
' сохраняем фон на новом участке
lngRtn = BitBlt(frmMon.picBack(Personag + 1).hDC, 0, 0, 50, 75, OffSp.hDC, x, y, vbSrcCopy)
' применяем маску
lngRtn = BitBlt(OffSp.hDC, x, y, 50, 75, frmMon.MonNeg.hDC, 0, 0, vbSrcAnd)
' рисуем
lngRtn = BitBlt(OffSp.hDC, x, y, 50, 75, frmMon.Mon.hDC, 0, 0, vbSrcPaint)
End Sub

Private Sub DoMerrir()
Dim Work As Long
    'ссиметричный рисунок позитива
    For y = 0 To lngH
    For x = 1 To lngW \ 2
    Work = frmMon.Mon.Point(x, y)
    frmMon.Mon.PSet (x, y), frmMon.Mon.Point(lngW - x, y)
    frmMon.Mon.PSet (lngW - x, y), Work
    Next x
    Next y
    'ссиметричный рисунок негатива
    For y = 0 To lngH
    For x = 1 To lngW \ 2
    Work = frmMon.MonNeg.Point(x, y)
    frmMon.MonNeg.PSet (x, y), frmMon.MonNeg.Point(lngW - x, y)
    frmMon.MonNeg.PSet (lngW - x, y), Work
    Next x
    Next y
    'доб в набор позитев
    frmMon.Mon.Picture = frmMon.Mon.Image
    Set imgX = frmMon.imlMon.ListImages.Add(, , frmMon.Mon.Picture)
    'доб в набор  негатив
    frmMon.MonNeg.Picture = frmMon.MonNeg.Image
    Set imgX = frmMon.imlMon.ListImages.Add(, , frmMon.MonNeg.Picture)
End Sub


