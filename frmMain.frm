VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Main"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   FillColor       =   &H00011BFC&
   ForeColor       =   &H0000DF00&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMusic 
      Interval        =   500
      Left            =   360
      Top             =   240
   End
   Begin VB.Label lblSave 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Сохранить"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000020FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Label lblMap 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Редактор карт"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000020FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   6600
      Width           =   2760
   End
   Begin VB.Label lblLoadGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Загрузить игру"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000020FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   5880
      Width           =   2760
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000020FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label lblNastr 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Опции"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000020FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label lblGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Играть"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000020FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   4440
      Width           =   1380
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function mciExecute _
Lib "winmm.dll" ( _
ByVal IpstrCommand As String _
) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim X As Long

Private Sub Form_Initialize()
    Member.MServiseForm = frmMain
    Member.GameDirectory = App.Path
    Load.Load_GMotion
    Company = "Office_Space"
    Level = 1
    Lifes = 3
    Money = 99
End Sub

Private Sub lblGame_Click()
    If lblGame.Caption = "Продолжить" Then
        'запуск игры
        SEvents.Game_Start True
        OffSp.Show
        Unload frmMain
    Else
        Level = 1
        Company = "Office_Space"
        Unload frmMain
        OffSp.Load_Game
    End If
End Sub

Private Sub lblNastr_Click()
    frmOpt.Show vbModal
End Sub

Private Sub lblSave_Click()
    SaveMode = True
    frmMain.Hide
    frmLoadGame.Show
    frmLoadGame.lblNew.Visible = True
    frmLoadGame.txtNew.Visible = True
End Sub

Private Sub lblLoadGame_Click()
    SaveMode = False
    frmLoadGame.Show
    frmLoadGame.lblNew.Visible = False
    frmLoadGame.txtNew.Visible = False
    frmMain.Hide
End Sub

Private Sub lblMap_Click()
    If lblGame.Caption = "Продолжить" Then
            Unload OffSp
            lblGame.Caption = "Играть"
    End If
    Member.MServiseForm = frmMain
    Map.Map_Editor
    frmMain.Hide
End Sub

Private Sub lblExit_Click()
    If lblGame.Caption = "Продолжить" Then
        Unload OffSp: lblGame.Caption = "Играть"
        ' показываем курсор мыши
        X = ShowCursor(True)
    Else
        Unload frmMain
        End
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblGame.Caption = "Продолжить" Then
        lblSave.Visible = True
    Else
        lblSave.Visible = False
    End If
    X = X - 8760
    Y = Y - 4440
    If X < 0 Or Y < 0 Then Exit Sub
    Y = Y \ 700
    lblGame.ForeColor = frmMain.FillColor: lblGame.Font.Size = 16
    lblSave.ForeColor = frmMain.FillColor: lblSave.Font.Size = 16
    lblLoadGame.ForeColor = frmMain.FillColor: lblLoadGame.Font.Size = 16
    lblMap.ForeColor = frmMain.FillColor: lblMap.Font.Size = 16
    lblNastr.ForeColor = frmMain.FillColor: lblNastr.Font.Size = 16
    lblExit.ForeColor = frmMain.FillColor: lblExit.Font.Size = 16
    Select Case Y
        Case 0:  lblGame.ForeColor = frmMain.ForeColor: lblGame.Font.Size = 18
        Case 1:  lblSave.ForeColor = frmMain.ForeColor: lblSave.Font.Size = 18
        Case 2:  lblLoadGame.ForeColor = frmMain.ForeColor: lblLoadGame.Font.Size = 18
        Case 3: lblMap.ForeColor = frmMain.ForeColor: lblMap.Font.Size = 18
        Case 4: lblNastr.ForeColor = frmMain.ForeColor: lblNastr.Font.Size = 18
        Case 5: lblExit.ForeColor = frmMain.ForeColor: lblExit.Font.Size = 18
    End Select
End Sub

