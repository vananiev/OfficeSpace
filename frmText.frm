VERSION 5.00
Begin VB.Form frmText 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   FillColor       =   &H000020FF&
   ForeColor       =   &H00000000&
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmText.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblDay 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000CE00&
      Height          =   255
      Left            =   600
      MouseIcon       =   "frmText.frx":0316
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblText 
      BackColor       =   &H00000000&
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
      Height          =   3975
      Left            =   120
      MouseIcon       =   "frmText.frx":0620
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Load()
    lblDay = "Δενό " & Level
    frmText.Height = (32 + (Len(lblText) \ 50 + 1) * 17) * 15
End Sub

Private Sub lblText_Click()
    Unload frmText
End Sub
