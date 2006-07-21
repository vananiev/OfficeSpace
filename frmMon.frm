VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMon 
   AutoRedraw      =   -1  'True
   Caption         =   "GMotion"
   ClientHeight    =   3810
   ClientLeft      =   180
   ClientTop       =   1350
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlFurn 
      Left            =   6120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox FurnNeg 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.PictureBox Furn 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2880
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Timer tmrWork 
      Interval        =   100
      Left            =   4680
      Top             =   840
   End
   Begin MSComctlLib.ImageList imlMon 
      Left            =   6120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox MonNeg 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   2160
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Mon 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1200
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mons As New Motion
Dim Map As New Map
Dim intCount As Integer

Public Sub Load()
    frmWork.Show
    frmWork.SetFocus
    ' Загрузка интерьера
    frmWork.ResetCaption "Загрузка интерьера"
    For intCount = 0 To 45
        frmWork.lblPer = Str(Int(intCount / 45 * 100)) & "%"
        Map.AddFurniture intCount, "Furniture"
    Next intCount
    Open GameDir & "\Image\Furniture\FurniturePlace.inform" For Binary As #1
        Get #1, , FurnW()
        Get #1, , FurnH()
    Close #1
    Unload frmWork
End Sub

Private Sub Form_load()
    Mon.Width = 50
    Mon.Height = 75
    MonNeg.Width = 50
    MonNeg.Height = 75
    Furn.Height = 110
    Furn.Width = 96
    FurnNeg.Height = 110
    FurnNeg.Width = 96
End Sub

