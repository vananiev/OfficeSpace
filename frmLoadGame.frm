VERSION 5.00
Begin VB.Form frmLoadGame 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Load"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   FillColor       =   &H000020FF&
   ForeColor       =   &H0000DF00&
   Icon            =   "frmLoadGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frGame 
      BackColor       =   &H00000000&
      Caption         =   "Загрузка"
      ForeColor       =   &H0036FE4A&
      Height          =   2775
      Left            =   3240
      TabIndex        =   0
      Top             =   3360
      Width           =   4695
      Begin VB.ListBox lstGame 
         BackColor       =   &H00000000&
         ForeColor       =   &H000020FF&
         Height          =   1425
         ItemData        =   "frmLoadGame.frx":000C
         Left            =   1920
         List            =   "frmLoadGame.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtNew 
         BackColor       =   &H00000000&
         ForeColor       =   &H0036FE4A&
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
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
         Left            =   720
         TabIndex        =   6
         Top             =   1080
         Width           =   1020
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
         Left            =   720
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblExit 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Отмена"
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
         Left            =   720
         TabIndex        =   4
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label lblNew 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Новая"
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
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmLoadGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intCount As Integer

Private Sub Form_Load()
    If SaveMode Then
        frGame.Caption = "Сохранение"
    Else
        frGame.Caption = "Загрузить"
    End If
    ListGame
End Sub

Private Sub lblDel_Click()
    lblLoad.ForeColor = frmMain.ForeColor
    Kill App.Path & "\Save Games\" & lstGame.List(lstGame.ListIndex)
    ListGame
    lblLoad.ForeColor = frmMain.FillColor
End Sub

Private Sub lblExit_Click()
    Unload frmLoadGame
    frmMain.Show
End Sub

Public Sub lblLoad_Click()
    Dim strFile As String
    Dim strSpFile() As String
    If SaveMode Then
        Open App.Path & "\Save Games\" & lstGame.List(lstGame.ListIndex) For Binary As #1
            Put #1, , " " & Company & " " & Level & " " & startLifes & " " & startMoney
        Close #1
    Else
        Open App.Path & "\Save Games\" & lstGame.List(lstGame.ListIndex) For Binary As #1
            strFile = Space(LOF(1))
            Get #1, , strFile
        Close #1
        strSpFile = Split(strFile, " ")
        Company = strSpFile(1)
        Level = Val(strSpFile(2))
        Lifes = Val(strSpFile(3))
        Money = Val(strSpFile(4))
        Unload frmLoadGame
        OffSp.Load_Game
        OffSp.Show
        lblExit.Visible = True
        Unload frmMain
    End If
End Sub

Private Sub ListGame()
Dim strfIleName As String
    lstGame.Clear
    ' сначала перечисляем все обычные файлы в текущем каталоге
    strfIleName = Dir(App.Path & "\Save Games\")
    Do While strfIleName <> ""
    lstGame.AddItem strfIleName
    intCount = intCount + 1
    strfIleName = Dir
    Loop
    If lstGame.ListCount <> 0 Then lstGame.Selected(0) = True
End Sub

Private Sub lblNew_Click()
        Open App.Path & "\Save Games\" & txtNew & ".svg" For Binary As #1
            Put #1, , " " & Company & " " & Level & " " & Lifes & " " & Money
        Close #1
End Sub
