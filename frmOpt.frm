VERSION 5.00
Begin VB.Form frmOpt 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   FillColor       =   &H000020FF&
   ForeColor       =   &H0000CE00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frOpt 
      BackColor       =   &H00000000&
      Caption         =   "Options"
      ForeColor       =   &H000020FF&
      Height          =   1575
      Left            =   4440
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H008080FF&
         Caption         =   "Ok"
         Height          =   255
         Left            =   1560
         MaskColor       =   &H000020FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtFps 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblFps 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fps:"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000020FF&
         Height          =   270
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    intFps = Val(txtFps)
    If intFps > 24 Or intFps < 1 Then intFps = 24
    Unload frmOpt
End Sub

Private Sub Form_Load()
    txtFps = intFps
End Sub
