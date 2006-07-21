VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   FillColor       =   &H000020FF&
   ForeColor       =   &H0000DF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frGame 
      BackColor       =   &H00000000&
      Caption         =   "Загрузить"
      ForeColor       =   &H0036FE4A&
      Height          =   2655
      Left            =   2640
      TabIndex        =   0
      Top             =   3120
      Width           =   6855
      Begin VB.TextBox txtNumFurn 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H000020FF&
         Height          =   195
         Left            =   6120
         TabIndex        =   8
         Text            =   "45"
         Top             =   2280
         Width           =   375
      End
      Begin VB.ListBox lstCompany 
         BackColor       =   &H00000000&
         ForeColor       =   &H000020FF&
         Height          =   1620
         ItemData        =   "frmLoad.frx":0000
         Left            =   2040
         List            =   "frmLoad.frx":0007
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox lstLevel 
         BackColor       =   &H00000000&
         ForeColor       =   &H000020FF&
         Height          =   1620
         ItemData        =   "frmLoad.frx":0017
         Left            =   4320
         List            =   "frmLoad.frx":001E
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblFurn 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Количество мебели"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   165
         Left            =   4320
         TabIndex        =   7
         Top             =   2280
         Width           =   1620
      End
      Begin VB.Label lblDel 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Удалить"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblLoad 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Загрузить"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lblNew 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Добавить уровень"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label lblExit 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Отменить"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intCnt As Integer
Dim intIndex As Integer
Dim imgX As ListImage

Private Sub Form_load()
    Set objReturn = ServiseForm
    Set ServiseForm = MapEditor
    txtNumFurn = maxFurn
    ListGame
End Sub

Private Sub lblLoad_Click()
    Dim strPersonag As String
    Dim strSpFurn() As String
    Dim strSpPersonag() As String
    Dim intN As Integer
    lblLoad.ForeColor = frmLoad.ForeColor
    MapEditor.txtCompany = Left(lstCompany.List(lstCompany.ListIndex), Len(lstCompany.List(lstCompany.ListIndex)) - 5)
    MapEditor.txtLevel = lstLevel.List(lstLevel.ListIndex)
    If Val(txtNumFurn) > 223 Or Val(txtNumFurn) < 45 Then txtNumFurn = 45
        'добавление мебели
    For intCnt = maxFurn To Val(txtNumFurn)
        Map.AddFurniture intCnt, "Furniture"
    Next intCnt
    'установка нового потолка
    If Val(txtNumFurn) > maxFurn Then maxFurn = Val(txtNumFurn)
    ' загрузка уровня
    Load.Load_Game MapEditor.txtCompany, lstLevel.ListIndex + 1, MapEditor.NumPersonag
    Mons.Fps
    MapEditor.imlFon.ListImages.Clear
    Set imgX = MapEditor.imlFon.ListImages.Add(, , MapEditor.Image)
    'Загрузка персонажей
    frmWork.ResetCaption "Загрузка персонажей"
    Load.ClsPicture
    For intN = 0 To 7
        frmWork.lblPer = Str(Int(intN / 15 * 100)) & "%"
        Mons.AddMons intN, "Personag"
    Next intN
    'определение количества персонажей и установка параметров
    Open GameDir & "\Companies\Levels\" & MapEditor.txtLevel & ".per" For Binary As #1
       strPersonag = Space(LOF(1))
       Get #1, , strPersonag
    Close #1
    MapEditor.strNumPersonag = strPersonag
    strSpPersonag = Split(strPersonag, " ")
    If UBound(strSpPersonag) < 0 Then
        MapEditor.NumPersonag = 0
    Else
        MapEditor.NumPersonag = UBound(strSpPersonag)
    End If
    MapEditor.Load_MapEditor
    MapEditor.SaveGame = True
    MapEditor.Show
    MapEditor.SetFocus
    Unload frmLoad
End Sub

Private Sub lblDel_Click()
Dim strCompany As String
    Open GameDir & "\Companies\" & lstCompany.List(lstCompany.ListIndex) For Binary As #1
        strCompany = Space(LOF(1))
        Get #1, , strCompany
    Close #1
    strCompany = Replace(strCompany, " " & lstLevel.List(lstLevel.ListIndex), "")
    Kill GameDir & "\Companies\" & lstCompany.List(lstCompany.ListIndex)
    Open GameDir & "\Companies\" & lstCompany.List(lstCompany.ListIndex) For Binary As #1
        Put #1, , strCompany
    Close #1
    On Error Resume Next
    Kill GameDir & "\Companies\levels\" & lstLevel.List(lstLevel.ListIndex) & ".per"
    On Error Resume Next
    Kill GameDir & "\Companies\levels\" & lstLevel.List(lstLevel.ListIndex) & ".plc"
    On Error Resume Next
    Kill GameDir & "\Companies\levels\" & lstLevel.List(lstLevel.ListIndex) & ".tip"
    intIndex = 0
    ListGame
End Sub

Private Sub lblNew_Click()
    Dim intN As Integer
    Dim intM As Integer
    Dim intNumer As Integer
    lblNew.ForeColor = frmLoad.ForeColor
    If Val(txtNumFurn) > 223 Or Val(txtNumFurn) < 45 Then txtNumFurn = 45
        ' Загрузка персонажей
    frmWork.ResetCaption "Загрузка персонажей"
    Load.ClsPicture
    For intN = 0 To 7
        frmWork.lblPer = Str(Int(intN / 15 * 100)) & "%"
        Mons.AddMons intN, "Personag"
    Next intN
        ' Загрузка интерьера
    frmWork.ResetCaption "Загрузка интерьера"
    'добавление мебели
    For intCnt = maxFurn To Val(txtNumFurn)
        Map.AddFurniture intCnt, "Furniture"
    Next intCnt
    'установка нового потолка
    If Val(txtNumFurn) > maxFurn Then maxFurn = Val(txtNumFurn)
    ' начальное состояние оффиса
    For intN = 0 To 15
        For intM = 0 To 42
            Place(intN, intM, 0) = 255
        Next intM
    Next intN
    MapEditor.imlFon.ListImages.Clear
    Set imgX = MapEditor.imlFon.ListImages.Add(, , MapEditor.Image)
    MapEditor.Load_MapEditor
    MapEditor.txtCompany = Left(lstCompany.List(lstCompany.ListIndex), Len(lstCompany.List(lstCompany.ListIndex)) - 5)
    MapEditor.Show
    MapEditor.SetFocus
    Unload frmLoad
End Sub

Private Sub lblExit_Click()
    Set ServiseForm = objReturn
    ServiseForm.Show
    Unload frmLoad
End Sub

Private Sub ListGame()
Dim strFileName As String
Dim strCompany As String
Dim strSpCompany() As String
Dim intN As Integer
    lstCompany.Clear
    ' сначала перечисляем все обычные файлы в текущем каталоге
    strFileName = Dir(GameDir & "\Companies\")
    Do While strFileName <> ""
        lstCompany.AddItem strFileName
        strFileName = Dir
    Loop
    If lstCompany.ListCount = 0 Then lstLevel.Clear: Exit Sub
    lstCompany.Selected(intIndex) = True
    lstLevel.Clear
    Open GameDir & "\Companies\" & lstCompany.List(lstCompany.ListIndex) For Binary As #1
        strCompany = Space(LOF(1))
        Get #1, , strCompany
    Close #1
    strSpCompany = Split(strCompany, " ")
    For intN = 1 To UBound(strSpCompany)
        lstLevel.AddItem strSpCompany(intN)
    Next intN
    If lstLevel.ListCount = 0 Then Exit Sub
    lstLevel.Selected(0) = True
End Sub

Private Sub lstCompany_DblClick()
    intIndex = lstCompany.ListIndex
    ListGame
End Sub

