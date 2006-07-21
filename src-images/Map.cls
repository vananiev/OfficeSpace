VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal x As Long, ByVal у As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long
Dim imgX As ListImage
Dim lngW As Long, lngH As Long
Dim x As Long, y As Long

Public Sub AddFurniture(NumPic As Integer, Width As Integer, Height As Integer, Name As String)
    FurnW(NumPic) = Width
    FurnH(NumPic) = Height
    'загрузка картиноки и сохранение в imlPon                                    '-16 временно
    frmMon.Furn = LoadPicture(App.Path & "\" & Replace(Name & "_" & Str(NumPic) & ".gif", " ", ""))
    frmMon.Furn.Refresh
    lngW = frmMon.Furn.ScaleWidth
    lngH = frmMon.Furn.ScaleHeight
    frmMon.Furn.ScaleWidth = lngW
    frmMon.Furn.ScaleHeight = lngH
    frmMon.FurnNeg.ScaleWidth = lngW
    frmMon.FurnNeg.ScaleHeight = lngH
    frmMon.Furn.Refresh
    frmMon.FurnNeg.Refresh
    For x = 0 To lngW
    For y = 0 To lngH
    'негатив
    If frmMon.Furn.Point(x, y) = 16777215 Then
        frmMon.FurnNeg.PSet (x, y), 16777215
    Else
        frmMon.FurnNeg.PSet (x, y), 0
    End If
    Next y
    Next x
    frmMon.FurnNeg.Refresh
    'позитив
    For x = 0 To lngW
    For y = 0 To lngH
        If frmMon.Furn.Point(x, y) = 16777215 Then frmMon.Furn.PSet (x, y), 0
    Next y
    Next x
    'доб в набор позитев
    frmMon.Furn.Picture = frmMon.Furn.Image
    Set imgX = frmMon.imlFurn.ListImages.Add(, , frmMon.Furn.Picture)
    'доб в набор  негатив
    frmMon.FurnNeg.Picture = frmMon.FurnNeg.Image
    Set imgX = frmMon.imlFurn.ListImages.Add(, , frmMon.FurnNeg.Picture)
End Sub

Public Sub FurnPaint(Xplace As Integer, Yplace As Integer, Position As Integer)
Dim lngRtn As Long
x = Xplace * 48 - (Yplace / 2 - Yplace \ 2) * 48 + 24
y = Yplace * 12 - 120
' прорисовка
frmMon.FurnNeg.Picture = frmMon.imlFurn.ListImages.Item(Position + 1).Picture
frmMon.Furn.Picture = frmMon.imlFurn.ListImages.Item(Position).Picture
' применяем маску
lngRtn = BitBlt(OffSp.hDC, x, y, 144, 168, frmMon.FurnNeg.hDC, 0, 0, vbSrcAnd)
' рисуем
lngRtn = BitBlt(OffSp.hDC, x, y, 144, 168, frmMon.Furn.hDC, 0, 0, vbSrcPaint)

End Sub

Public Sub DectoPlc(x As Single, y As Single, Xp As Single, Yp As Single)
    Yp = y \ 12
    Xp = (x - Abs((Yp / 2 - Yp \ 2) * 2 - 1) * 24) \ 48
End Sub

Public Sub PlstoDec(x As Single, y As Single, Xp As Single, Yp As Single)
    Xp = x * 48 - (y / 2 - y \ 2) * 48 + 24
    Yp = y * 12 + 12
End Sub


