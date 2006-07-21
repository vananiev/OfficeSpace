VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FerstLoad"
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
Dim lngRtn As Long
Dim lngW As Long, lngH As Long
Dim mX As Single, mY As Single

Public Sub AddNFurniture()
Dim strFileName As String
'создаем новый каталог
On Error Resume Next
MkDir (GameDir & "\Image\Furniture\NFurniture\")
' перечисляем все обычные файлы в текущем каталоге
strFileName = Dir(GameDir & "\Image\Furniture\")
Do While strFileName <> ""
    'загрузка картиноки и сохранение в imlPon
    On Error GoTo NLoop
    frmMon.Furn = LoadPicture(GameDir & "\Image\Furniture\" & strFileName)
    lngW = frmMon.Furn.ScaleWidth
    lngH = frmMon.Furn.ScaleHeight
    'негатив
    For mX = 0 To lngW
        For mY = 0 To lngH
            If frmMon.Furn.Point(mX, mY) = 0 Then
                frmMon.FurnNeg.PSet (mX, mY), 16777215
            Else
                frmMon.FurnNeg.PSet (mX, mY), 0
            End If
        Next mY
    Next mX
    'доб в набор  негатив
    frmMon.FurnNeg.Picture = frmMon.FurnNeg.Image
    SavePicture frmMon.FurnNeg.Picture, GameDir & "\Image\Furniture\NFurniture\" & strFileName
NLoop:
strFileName = Dir
Loop
End Sub

Public Sub AddNPersonag()
Dim mX As Single, mY As Single
Dim strFileName As String
'создаем новый каталог
On Error Resume Next
MkDir (GameDir & "\Image\Personag\NPersonag\")
On Error Resume Next
MkDir (GameDir & "\Image\Personag\MPersonag\")
On Error Resume Next
MkDir (GameDir & "\Image\Personag\NMPersonag\")
' перечисляем все обычные файлы в текущем каталоге
strFileName = Dir(GameDir & "\Image\Personag\")
Do While strFileName <> ""
    'загрузка картиноки и сохранение в imlPon
    frmMon.Mon.Picture = LoadPicture(GameDir & "\Image\Personag\" & strFileName)
    lngW = frmMon.Mon.ScaleWidth
    lngH = frmMon.Mon.ScaleHeight
    'негатив
    For mX = 0 To lngW
        For mY = 0 To lngH
            If frmMon.Mon.Point(mX, mY) = 0 Then
                frmMon.MonNeg.PSet (mX, mY), 16777215
            Else
                frmMon.MonNeg.PSet (mX, mY), 0
            End If
        Next mY
    Next mX
    'доб в набор  негатив
    frmMon.MonNeg.Picture = frmMon.MonNeg.Image
    SavePicture frmMon.MonNeg.Picture, GameDir & "\Image\Personag\NPersonag\" & strFileName
    DoMerrir strFileName
    strFileName = Dir
Loop
End Sub

Private Sub DoMerrir(strFileName As String)
    'ссиметричный рисунок позитива
    frmMon.BackColor = vbBlack
    frmMon.Cls
    For mX = 0 To lngW - 1
        lngRtn = BitBlt(frmMon.hDC, mX, 0, 1, lngH, frmMon.Mon.hDC, lngW - mX - 1, 0, vbSrcPaint)
    Next mX
    lngRtn = BitBlt(frmMon.Mon.hDC, 0, 0, lngW, lngH, frmMon.hDC, 0, 0, vbSrcCopy)
    frmMon.Cls
    'ссиметричный рисунок негатива
    For mX = 0 To lngW - 1
        lngRtn = BitBlt(frmMon.hDC, mX, 0, 1, lngH, frmMon.MonNeg.hDC, lngW - mX - 1, 0, vbSrcPaint)
    Next mX
    lngRtn = BitBlt(frmMon.MonNeg.hDC, 0, 0, lngW, lngH, frmMon.hDC, 0, 0, vbSrcCopy)
    frmMon.Cls
    'доб в набор позитев
    frmMon.Mon.Picture = frmMon.Mon.Image
    SavePicture frmMon.Mon.Picture, GameDir & "\Image\Personag\MPersonag\" & strFileName
    'доб в набор  негатив
    frmMon.MonNeg.Picture = frmMon.MonNeg.Image
    SavePicture frmMon.MonNeg.Picture, GameDir & "\Image\Personag\NMPersonag\" & strFileName
End Sub

