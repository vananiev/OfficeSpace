VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim X As Long
Dim intCnt As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Boolean) As Long
Private Declare Function mciExecute _
Lib "winmm.dll" ( _
ByVal IpstrCommand As String _
) As Long

Public Sub GameLoad()
    Dim strCompany As String
    Dim strSpComp() As String
    Dim strSMS As String
    Dim NumFurn As Integer
    Member.MServiseForm = OffSp
    On Error GoTo WinGame
    Load.Load_Game Company, Level, NumPersonag
    If Level = 1 Then
        Lifes = 3: Money = 173
        startLifes = Lifes
        startMoney = Money
    End If
    'поиск папок
    Files = 0
    For Xp = 0 To 15
        For Yp = 6 To 42
            NumFurn = Member.MPlace(Xp, Yp, 0)
            If NumFurn = 28 Then Files = Files + 1
        Next Yp
    Next Xp
    'загрузка иконок
    IconPaint
    OffSp.Show
    ' Информация дня
    Open App.Path & "\companies\" & Company & ".comp" For Binary As #1
        strCompany = Space(LOF(1))
        Get #1, , strCompany
    Close #1
    strSpComp = Split(strCompany, " ")
    On Error GoTo RunGame
    Open App.Path & "\Companies\Levels\" & strSpComp(Level) & ".tip" For Input As #1
        Line Input #1, strSMS
    Close #1
    frmText.lblText = strSMS
    frmText.Load
    frmText.Show vbModal
RunGame:
    'запуск игры
    Game_Start True
    Exit Sub
WinGame:
    Level = Level - 1
    mciExecute "Play C:\Sound\Win.wav"
    frmText.lblText = "Поздравляем вы дожили до счастливого летнего отпуска"
    frmText.Load
    frmText.Show vbModal
    Unload OffSp
    frmMain.Show
End Sub

Public Sub Events(Bonus As Byte)
Dim Result As Long
    'события
    If Money <= 0 Then
        SEvents.Game_Start False
        mciExecute "Play C:\Sound\End.wav"
        frmText.lblText = "У вас нет денег, вывозвращаетесь в начало месяца"
        frmText.Load
        frmText.Show vbModal
        Level = Level \ 5 * 5
        OffSp.Load_Game
    End If
    If Files = 0 Then
        SEvents.Game_Start False
        mciExecute "Play C:\Sound\Win.wav"
        frmText.lblText = "Вы прошли уровень"
        frmText.Load
        frmText.Show vbModal
        Level = Level + 1
        startLifes = Lifes
        startMoney = Money
        OffSp.Load_Game
    End If
    If Lifes <= 0 Then
        SEvents.Game_Start False
        mciExecute "Play C:\Sound\End.wav"
        frmText.lblText = "Вас выгнали с работы"
        frmText.Load
        frmText.Show vbModal
        Unload OffSp
        frmMain.Show
    End If
    ' пополнение бонусов
    Select Case Bonus
        Case 28: Files = Files - 1: IconPaint: mciExecute "Play C:\Sound\Files.wav"
        Case 27: Lifes = Lifes + 1: IconPaint: mciExecute "Play C:\Sound\Heard.wav"
        Case 29: Randomize Timer: Money = Money + 50 + Int(Rnd * 26): IconPaint: mciExecute "Play C:\Sound\Money.wav"
End Select
End Sub

Public Sub IconPaint()
    For intCnt = 0 To OffSp.imgFile.Count - 1
        OffSp.imgFile(intCnt).Picture = Nothing
    Next intCnt
    For intCnt = 0 To OffSp.imgLife.Count - 1
        OffSp.imgLife(intCnt).Picture = Nothing
    Next intCnt
    OffSp.lblNumFiles = Files
    For intCnt = 0 To Files - 1
        On Error Resume Next
        OffSp.imgFile(intCnt).Picture = LoadPicture(App.Path & "\Image\More\File.gif")
    Next intCnt
    OffSp.lblNumLifes = Lifes
    For intCnt = 0 To Lifes - 1
        On Error Resume Next
        OffSp.imgLife(intCnt).Picture = LoadPicture(App.Path & "\Image\More\Lifes.gif")
    Next intCnt
    OffSp.lblNumMoney = Abs(Money)
    If Money < 0 Then
        OffSp.lblPay.Visible = True
    Else
        OffSp.lblPay.Visible = False
    End If
End Sub

Public Function Game_Start(Direction As Boolean)
    OffSp.tmrMove.Enabled = Direction
    OffSp.tmrMons.Enabled = Direction
    OffSp.tmrFps.Enabled = Direction
    ' курсор мыши
    X = ShowCursor(Not (Direction))
End Function

