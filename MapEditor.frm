VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MapEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FEBD52&
   BorderStyle     =   0  'None
   Caption         =   "MapEditor"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlFon 
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox ButtonLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   8280
      Width           =   495
   End
   Begin VB.PictureBox picFon 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   7665
      Width           =   12000
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Exit"
         Default         =   -1  'True
         Height          =   255
         Left            =   11040
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "¬ыход"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtCompany 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00011BFC&
         Height          =   285
         Left            =   10920
         TabIndex        =   9
         Text            =   "New"
         ToolTipText     =   "¬ведите им€ компании"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtLevel 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00011BFC&
         Height          =   285
         Left            =   10920
         TabIndex        =   6
         Text            =   "New"
         ToolTipText     =   "¬ведите им€ уровн€"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Ok 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Generate"
         Height          =   255
         Left            =   11040
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "—оздать карту"
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.PictureBox ButtonRight 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   9000
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox ButtonRight 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   6240
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox ButtonLeft 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   6840
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblAll 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00011BFC&
         Height          =   165
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label lblPers 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Personag"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00011BFC&
         Height          =   165
         Left            =   6960
         TabIndex        =   8
         Top             =   120
         Width           =   675
      End
      Begin VB.Label lblFurn 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Furniture"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   6
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00011BFC&
         Height          =   165
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   675
      End
      Begin VB.Image imgPersonag 
         Height          =   735
         Index           =   2
         Left            =   8520
         Stretch         =   -1  'True
         Top             =   360
         Width           =   525
      End
      Begin VB.Image imgPersonag 
         Height          =   735
         Index           =   1
         Left            =   7920
         Stretch         =   -1  'True
         Top             =   360
         Width           =   525
      End
      Begin VB.Image imgPersonag 
         Height          =   735
         Index           =   0
         Left            =   7320
         Stretch         =   -1  'True
         Top             =   360
         Width           =   525
      End
      Begin VB.Image imgFurniture 
         Height          =   735
         Index           =   7
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgFurniture 
         Height          =   735
         Index           =   6
         Left            =   4920
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgFurniture 
         Height          =   735
         Index           =   5
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgFurniture 
         Height          =   735
         Index           =   4
         Left            =   3480
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgFurniture 
         Height          =   735
         Index           =   3
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgFurniture 
         Height          =   735
         Index           =   2
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgFurniture 
         Height          =   735
         Index           =   1
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image imgFurniture 
         Height          =   735
         Index           =   0
         Left            =   600
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "MapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intFurnStart As Integer                      ' начало иконок мебели
Dim intPersonagStart As Integer                  ' начало иконок персонажей
Dim FurnPaint As Boolean                         ' флаг прорисовки мебели/персонажей
Dim intPicture As Integer                        ' картинка мебели/персонажей
Public NumPersonag As Integer                       ' последовательность загрузки персонажей
Public strNumPersonag  As String                    ' последовательность загрузки персонажей (в файлах)
Public SaveGame As Boolean                          ' режим загруженной/новой игры
Dim blnLoad  As Boolean
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
Dim imgX As ListImage
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Sub Load_MapEditor()
    Dim m As Integer
    Dim X As Long
    ' восстанавливаем курсор мыши
    X = ShowCursor(True)
    Dim imgX As ListImage
    ' загрузка кнопок
    ButtonLeft(0).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_0.gif")
    ButtonLeft(1).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_0.gif")
    ButtonRight(0).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_0.gif")
    ButtonRight(1).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_0.gif")
    IconsPaint
    FurnPaint = True
    intPicture = 1
    blnLoad = True
End Sub

Private Sub ButtonRight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonRight(Index).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_1.gif")
End Sub

Private Sub ButtonLeft_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonLeft(Index).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_1.gif")
End Sub

Private Sub ButtonRight_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonRight(0).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_0.gif")
    ButtonRight(1).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_0.gif")
    If Not (blnLoad) Then Exit Sub
    If Index = 0 Then
        intFurnStart = intFurnStart + 1
        If (frmMon.imlFurn.ListImages.Count) / 2 - 7 <= intFurnStart + 1 Then intFurnStart = intFurnStart - 1
    Else
        intPersonagStart = intPersonagStart + 1
        If (frmMon.imlMon.ListImages.Count) / 48 - 4 <= intPersonagStart Then intPersonagStart = intPersonagStart - 1
    End If
IconsPaint
End Sub

Private Sub ButtonLeft_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonLeft(0).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_0.gif")
    ButtonLeft(1).Picture = LoadPicture(GameDir & "\Image\More\" & "Button_0.gif")
    If Not (blnLoad) Then Exit Sub
    If Index = 0 Then
        intFurnStart = intFurnStart - 1
        If 0 > intFurnStart Then intFurnStart = intFurnStart + 1
    Else
        intPersonagStart = intPersonagStart - 1
        If 0 > intPersonagStart Then intPersonagStart = intPersonagStart + 1
    End If
IconsPaint
End Sub

Private Sub IconsPaint()
    For intCount = 0 To imgFurniture.Count
        On Error Resume Next
        imgFurniture(intCount).Picture = frmMon.imlFurn.ListImages.Item((intCount + intFurnStart) * 2 + 1).Picture
    Next intCount
    For intCount = 0 To imgPersonag.Count
        On Error Resume Next
        imgPersonag(intCount).Picture = frmMon.imlMon.ListImages.Item((intCount + intPersonagStart) * 48 + 9).Picture
    Next intCount
End Sub

Private Sub cmdExit_Click()
    Set ServiseForm = objReturn
    ServiseForm.Show
    Unload MapEditor
End Sub

Private Sub imgFurniture_Click(Index As Integer)
    If Not (blnLoad) Then Exit Sub
    If imgFurniture(Index).Picture = 0 Then Exit Sub
    FurnPaint = True
    intPicture = (intFurnStart + Index) * 2 + 1
    If (intPicture - 1) / 2 < 11 Then
        lblAll.Visible = True
    Else
        lblAll.Visible = False
    End If
End Sub

Private Sub imgPersonag_Click(Index As Integer)
    If Not (blnLoad) Then Exit Sub
    If imgPersonag(Index).Picture = 0 Then Exit Sub
    FurnPaint = False
    intPicture = 9
    Personag = Index + intPersonagStart
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (blnLoad) Then Exit Sub
    MapEditor.Picture = imlFon.ListImages(1).Picture
    If Y > 500 Or Y < 70 Then Exit Sub
    Map.DectoPlc X, Y, Xp, Yp
    If FurnPaint Then
        Map.FurnPaint Xp, Yp, intPicture
    Else
        Map.PlstoDec Xp, Yp, X, Y
        Mons.Paint Personag, X - 25, Y - 55, intPicture
    End If
    Refresh
    If Button = 1 Then Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (blnLoad) Then Exit Sub
    If Y > 500 Then Exit Sub
    Map.DectoPlc X, Y, Xp, Yp
    If FurnPaint Then
        If (intPicture - 1) / 2 < 11 Then
            Place(Xp, Yp, 3) = (intPicture - 1) / 2 + 16
        Else
            Load.PlaceMassiv (intPicture - 1) / 2, Xp, Yp
        End If
    Else
        If NumPersonag = 16 Then MsgBox "Ѕольшее количество персонажей не допускаетс€", vbInformation, "Information": Exit Sub
        Place(Xp, Yp, 0) = NumPersonag
        strNumPersonag = strNumPersonag & " " & Personag
        NumPersonag = NumPersonag + 1
    End If
    imlFon.ListImages.Clear
    Set imgX = imlFon.ListImages.Add(, , MapEditor.Image)
End Sub

Private Function Dublicat(strMain As String, strSeach As String, Optional Work = False) As Integer
Dim strSpMain() As String
    strSpMain = Split(strMain, " ")
    For intCount = 1 To UBound(strSpMain)
        If Work Then
            Dublicat = Dublicat(strMain, strSpMain(Val(strSeach) - 15), False)
            Exit Function
        Else
            If strSpMain(intCount) = strSeach Then Dublicat = intCount: Exit Function
        End If
    Next intCount
    Dublicat = 0
End Function

Private Sub lblAll_Click()
     ' начальное состо€ние оффиса
    For Yp = 6 To 41
        For Xp = 15 To 0 Step (-1)
            Place(Xp, Yp, 3) = (intPicture - 1) / 2 + 16
            Map.FurnPaint Xp, Yp, intPicture
        Next Xp
    Next Yp
    imlFon.ListImages.Clear
    Set imgX = imlFon.ListImages.Add(, , MapEditor.Image)
End Sub

Private Sub Ok_Click()
    If Not (blnLoad) Then Exit Sub
    Dim strCompany As String
    Dim Numer As Integer
    Dim intOverwrite As Integer
    Dim strFileName As String
        ' проверка существовани€ уровней
        ' сначала перечисл€ем все обычные файлы в текущем каталоге
    strFileName = Dir(GameDir & "\Companies\Levels\")
    Do While strFileName <> ""
        If strFileName = txtLevel.Text & ".plc" Then
            intOverwrite = MsgBox("Level already exists. " & "Overwrite it?", vbYesNo)
            'если пользователь выбирает No, выходим из этой процедуры
            If intOverwrite = vbNo Then
                Close #1: txtLevel.SetFocus: Exit Sub
            End If
        End If
    strFileName = Dir
    Loop
    ' сохранение кампании
    Open GameDir & "\Companies\" & txtCompany & ".comp" For Binary As #1
        strCompany = Space(LOF(1))
        Get #1, , strCompany
        Numer = Dublicat(strCompany, txtLevel.Text)
        If Numer Then
            GoTo SaveLevel
        Else
            Put #1, , " " & txtLevel.Text
        End If
SaveLevel:
    Close #1
    'сохранение уровней
    On Error Resume Next
    Kill GameDir & "\Companies\Levels\" & txtLevel.Text & ".per"
    On Error Resume Next
    Kill GameDir & "\Companies\Levels\" & txtLevel.Text & ".plc"
    Open GameDir & "\Companies\Levels\" & txtLevel.Text & ".plc" For Binary As #1
        Put #1, , Place()
    Close #1
    Open GameDir & "\Companies\Levels\" & txtLevel.Text & ".per" For Binary As #1
        Put #1, , strNumPersonag
    Close #1
End Sub

