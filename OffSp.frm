VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OffSp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FED085&
   BorderStyle     =   0  'None
   Caption         =   "OfficeSpace"
   ClientHeight    =   8970
   ClientLeft      =   180
   ClientTop       =   1350
   ClientWidth     =   11970
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "OffSp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   0  'User
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMenu 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H0036FE4A&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   798
      TabIndex        =   0
      Top             =   7995
      Width           =   11970
      Begin VB.Label lblPay 
         AutoSize        =   -1  'True
         BackColor       =   &H0036FE4A&
         Caption         =   "долга"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   225
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0036FE4A&
         Caption         =   "рублей"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   225
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblMoney 
         BackColor       =   &H0036FE4A&
         Caption         =   "Money"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   960
      End
      Begin VB.Image imgLife 
         Height          =   375
         Index           =   4
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgLife 
         Height          =   375
         Index           =   3
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgFile 
         Height          =   375
         Index           =   6
         Left            =   11160
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgFile 
         Height          =   375
         Index           =   5
         Left            =   10680
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgFile 
         Height          =   375
         Index           =   4
         Left            =   10200
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgFile 
         Height          =   375
         Index           =   3
         Left            =   9720
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgFile 
         Height          =   375
         Index           =   2
         Left            =   9240
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgFile 
         Height          =   375
         Index           =   1
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgFile 
         Height          =   375
         Index           =   0
         Left            =   8280
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblNumMoney 
         BackColor       =   &H0036FE4A&
         Caption         =   "120"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblNumLifes 
         BackColor       =   &H0036FE4A&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   0
         Width           =   1290
      End
      Begin VB.Image imgLife 
         Height          =   375
         Index           =   2
         Left            =   5280
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgLife 
         Height          =   375
         Index           =   1
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgLife 
         Height          =   375
         Index           =   0
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblFiles 
         AutoSize        =   -1  'True
         BackColor       =   &H0036FE4A&
         Caption         =   "Files"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   375
         Left            =   8400
         TabIndex        =   3
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblNumFiles 
         BackColor       =   &H0036FE4A&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   375
         Left            =   9240
         TabIndex        =   2
         Top             =   0
         Width           =   810
      End
      Begin VB.Label lblLifes 
         AutoSize        =   -1  'True
         BackColor       =   &H0036FE4A&
         Caption         =   "Lifes"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   375
         Left            =   4440
         TabIndex        =   1
         Top             =   0
         Width           =   645
      End
   End
   Begin MSComctlLib.ImageList imlFon 
      Left            =   0
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer tmrFps 
      Enabled         =   0   'False
      Interval        =   42
      Left            =   0
      Top             =   4680
   End
   Begin VB.Timer tmrMons 
      Enabled         =   0   'False
      Interval        =   42
      Left            =   0
      Top             =   4200
   End
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
Dim Bonus As Byte
Dim intTimer As Integer
Dim intCnt As Integer
Dim X As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function mciExecute _
Lib "winmm.dll" ( _
ByVal IpstrCommand As String _
) As Long

Public Sub Load_Game()
    SEvents.GameLoad
    tmrFps.Interval = Int(1000 / intFps)
    ' курсор мыши
    X = ShowCursor(False)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        'останов игры
        SEvents.Game_Start False
        frmMain.lblGame.Caption = "Продолжить"
        frmMain.Show
    End If
    Mons.KeyPress 0, KeyCode, Shift
End Sub

Private Sub tmrFps_Timer()
    Mons.Fps
End Sub

Private Sub tmrMove_Timer()
    Bonus = Mons.PersonagMove(0)
    SEvents.Events Bonus
End Sub
Private Sub tmrMons_Timer()
    On Error Resume Next
    Mons.MonstrsMove NumPersonag - 1
    If intTimer > 0 Then intTimer = intTimer - 1
    If Member.MCatch Then
        If intTimer = 0 Then
            Lifes = Lifes - 1
            Money = Money - Int(Rnd * 51) - 50
            SEvents.IconPaint
            SEvents.Game_Start False
            mciExecute "Play C:\Sound\Heard.wav"
            frmText.lblText = "Вас поймал и оштрафовал начальник"
            frmText.Load
            frmText.Show vbModal
            SEvents.Game_Start True
            intTimer = 30
        End If
    End If
End Sub

Private Sub Form_Terminate()
    ' показываем курсор мыши
    X = ShowCursor(True)
End Sub
