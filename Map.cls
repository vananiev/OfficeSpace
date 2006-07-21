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
Attribute VB_Exposed = True
Option Explicit
Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal X As Long, ByVal у As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long
Dim imgX As ListImage
Dim X As Single, Y As Single
Dim lngRtn As Long

Public Sub AddFurniture(NumerPicture As Integer, Name As String)
DoEvents
        'загрузка картиноки и сохранение в imlPon
    On Error GoTo ExitSub
    frmMon.Furn = LoadPicture(GameDir & "\Image\Furniture\" & Name & "_" & NumerPicture & ".gif")
    On Error GoTo ExitSub
    frmMon.FurnNeg = LoadPicture(GameDir & "\Image\Furniture\NFurniture\" & Name & "_" & NumerPicture & ".gif")
    'доб в набор позитев
    Set imgX = frmMon.imlFurn.ListImages.Add(, , frmMon.Furn.Picture)
    'доб в набор  негатив
    Set imgX = frmMon.imlFurn.ListImages.Add(, , frmMon.FurnNeg.Picture)
ExitSub:
End Sub

Public Sub FurnPaint(Xplace As Integer, Yplace As Integer, Position As Integer)
Map.PlstoDec Xplace, Yplace, X, Y
X = X - 47
Y = Y - 95
' прорисовка
frmMon.FurnNeg.Picture = frmMon.imlFurn.ListImages.Item(Position + 1).Picture
frmMon.Furn.Picture = frmMon.imlFurn.ListImages.Item(Position).Picture
' применяем маску
lngRtn = BitBlt(ServiseForm.hDC, X, Y, 96, 110, frmMon.FurnNeg.hDC, 0, 0, vbSrcAnd)
' рисуем
lngRtn = BitBlt(ServiseForm.hDC, X, Y, 96, 110, frmMon.Furn.hDC, 0, 0, vbSrcPaint)
End Sub

Public Sub DectoPlc(X As Single, Y As Single, Xp As Integer, Yp As Integer)
    Yp = Y \ 12
    Xp = (X - Abs((Yp / 2 - Yp \ 2) * 2 - 1) * 24) \ 48
End Sub

Public Sub PlstoDec(Xp As Integer, Yp As Integer, X As Single, Y As Single)
    X = Xp * 48 + Abs((Yp / 2 - Yp \ 2) * 2 - 1) * 24 + 24
    Y = Yp * 12 + 10
End Sub
Public Sub Map_Editor()
    Dim intN As Integer
    Dim intM As Integer
    frmWork.Show
    frmWork.SetFocus
        'очистка рисунков
    Load.ClsPicture
    Unload frmWork
    frmLoad.Show
    frmLoad.SetFocus
End Sub
