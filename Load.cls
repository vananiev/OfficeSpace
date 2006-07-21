VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Load"
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

Public Sub Load_GMotion()
On Error GoTo FerstLoad
    Open GameDir & "\Image\Personag\Npersonag\Personag_0_0.gif" For Binary As #1
    Close #1
lReturn:
    frmMon.Load
    maxFurn = 45
    Exit Sub
FerstLoad:
    frmWork.ResetCaption "Инициализация игры..."
    frmWork.Show
    DoEvents
    FerstLoad.AddNFurniture
    FerstLoad.AddNPersonag
    Unload frmWork
    GoTo lReturn
End Sub

Public Sub Load_Game(Company As String, Level As Integer, NumPersonag As Integer)
    Dim imgX As ListImage
    Dim Xstr As Integer
    Dim Ystr As Integer
    Dim sX As Single
    Dim sY As Single
    ServiseForm.Picture = Nothing
    frmWork.ResetCaption "      Загрузка игры"
    frmWork.Show
    DoEvents
    'очистка рисунков мебели/персонажей
    ClsPicture
    'загрузка игры
    LoadGameFiles Company, Level, NumPersonag
    ' определение координат персонажей плитки
    For Yp = 0 To 42
        For Xp = 15 To 0 Step -1
            Personag = Place(Xp, Yp, 3)
            If Personag > 15 Then            ' поиск плитки
                Map.FurnPaint Xp, Yp, (Personag - 16) * 2 + 1
            End If
        Next Xp
    Next Yp
    For Yp = 0 To 42
         For Xp = 0 To 15
            Personag = Place(Xp, Yp, 0)
            If Personag > 31 And Personag < 255 Then        ' поиск мебели
                Xstr = Place(Xp, Yp, 1)
                Ystr = Place(Xp, Yp, 2)
                Map.FurnPaint Xstr, Ystr, (Personag - 16) * 2 + 1
            End If
            If Personag < 16 Then  ' поиск персонажей
                Map.PlstoDec Xp, Yp, X(Personag), Y(Personag)
                X(Personag) = X(Personag) - 23
                Y(Personag) = Y(Personag) - 57
            End If
        Next Xp
    Next Yp
    ServiseForm.imlFon.ListImages.Clear
    Set imgX = ServiseForm.imlFon.ListImages.Add(, , ServiseForm.Image)
    Unload frmWork
End Sub

Public Sub LoadFurniture(NumerFurniture As Integer, Xp As Integer, Yp As Integer, Optional NumerBanc = 0)
    ' номер интерьера (0<=x<=239)
    'загрузка интерьера
    PlaceMassiv NumerFurniture, Xp, Yp
    ' прорисовка интерьера
    Map.FurnPaint Xp, Yp, (NumerFurniture * 2) + 1
End Sub

Public Sub PlaceMassiv(NumerFurniture As Integer, Xp As Integer, Yp As Integer)
    Dim Xstp As Integer
    Dim Ystp As Integer
    Dim X As Single
    Dim Y As Single
    Dim intWork As Integer
    Dim Width As Integer, Height As Integer
    Dim N As Integer
    Xstp = Xp
    Ystp = Yp
    Width = FurnW(NumerFurniture)
    Height = FurnH(NumerFurniture)
    Map.PlstoDec Xstp, Ystp, X, Y
        'место установки самая нижняя кл. интерьера
        Place(Xp, Yp, 0) = NumerFurniture + 16  ' =интерьер
        Place(Xp, Yp, 1) = Xp
        Place(Xp, Yp, 2) = Yp
         ' up side
    For N = 1 To Height - 1
        Y = Y - 12
        X = X + 24
        Map.DectoPlc X, Y, Xstp, Ystp
        Place(Xstp, Ystp, 0) = NumerFurniture + 16  ' =интерьер
        Place(Xstp, Ystp, 1) = Xp
        Place(Xstp, Ystp, 2) = Yp
    Next N
        'right side
    For N = 1 To Width - 1
        Y = Y - 12
        X = X - 24
        Map.DectoPlc X, Y, Xstp, Ystp
        Place(Xstp, Ystp, 0) = NumerFurniture + 16  ' =интерьер
        Place(Xstp, Ystp, 1) = Xp
        Place(Xstp, Ystp, 2) = Yp
    Next N
        'doun side
    For N = 1 To Height - 1
        Y = Y + 12
        X = X - 24
        Map.DectoPlc X, Y, Xstp, Ystp
        Place(Xstp, Ystp, 0) = NumerFurniture + 16  ' =интерьер
        Place(Xstp, Ystp, 1) = Xp
        Place(Xstp, Ystp, 2) = Yp
    Next N
        'left side
    For N = 1 To Width - 1
        Y = Y + 12
        X = X + 24
        Map.DectoPlc X, Y, Xstp, Ystp
        Place(Xstp, Ystp, 0) = NumerFurniture + 16  ' =интерьер
        Place(Xstp, Ystp, 1) = Xp
        Place(Xstp, Ystp, 2) = Yp
    Next N
End Sub

Public Sub ClsPicture()
    Dim mintN As Integer
    Dim ListCount As Integer
    'очистка рисунков персонажей
    ListCount = frmMon.imlMon.ListImages.Count
    For mintN = ListCount To 1 Step -1
        frmMon.imlMon.ListImages.Remove (mintN)
    Next mintN
End Sub

Private Sub LoadGameFiles(Company As String, Level As Integer, NumPersonag As Integer)
Dim strPers As String
Dim strSpPers() As String
Dim strCompany As String
Dim strSpComp() As String
Dim Name As String
Dim mMaxFurn As Integer
Dim intN As Integer
    Open GameDir & "\Companies\" & Company & ".comp" For Binary As #1
        strCompany = Space(LOF(1))
        Get #1, , strCompany
    Close #1
    strSpComp = Split(strCompany, " ")
    Name = strSpComp(Level)
    Open GameDir & "\Companies\Levels\" & Name & ".plc" For Binary As #1
        Get #1, , Place()
    Close #1
    Open GameDir & "\Companies\Levels\" & Name & ".per" For Binary As #1
        strPers = Space(LOF(1))
        Get #1, , strPers
    Close #1
    strSpPers = Split(strPers, " ")
        ' загрузка персонажей
    NumPersonag = 0
    For intN = 1 To UBound(strSpPers)
        Mons.AddMons Int(Val(strSpPers(intN))), "Personag"
        intPos(intN - 1) = 17
        blnMove(intN - 1) = False
        strDir(intN - 1) = "stand"
        SeeDistance(intN - 1) = intN - 1
        NumPersonag = NumPersonag + 1
    Next intN
        ' загрузка мебели
    mMaxFurn = maxFurn
    For Yp = 0 To 42
        For Xp = 15 To 0 Step -1
            Personag = Place(Xp, Yp, 0)
            If Personag > mMaxFurn And Personag <> 255 Then ' поиск max рисунка мебели
                maxFurn = Personag
            End If
        Next Xp
    Next Yp
    For intCount = maxFurn To mMaxFurn
        Map.AddFurniture intCount, "Furniture"
    Next intCount
    'установка нового потолка
    If mMaxFurn > maxFurn Then maxFurn = mMaxFurn
End Sub
