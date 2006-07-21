VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMon 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Mon 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Left            =   2760
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.PictureBox MonNeg 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Index           =   1
      Left            =   600
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Timer tmrWork 
      Interval        =   100
      Left            =   0
      Top             =   720
   End
   Begin VB.PictureBox Furn 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   2760
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.PictureBox FurnNeg 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   0
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
   Begin MSComctlLib.ImageList imlFurn 
      Left            =   5640
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlMon 
      Left            =   5640
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
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

Private Sub Form_Load()
Mon.Width = 50
Mon.Height = 75
MonNeg.Width = 50
MonNeg.Height = 75
picBack(1).Height = 75
picBack(1).Width = 50
For intCount = 2 To NumPic + 1
Load picBack(intCount)
picBack(intCount).Height = 75
picBack(intCount).Width = 50
Next intCount
OffSp.Show
Mons.AddMon 0, "OfficeMan"
Mons.AddMon 1, "Shef"
frmWork.Show
frmWork.SetFocus
For intCount = 0 To 15
    Map.AddFurniture intCount, 1, 1, "Furniture"
    frmWork.lblPer = Str(intCount / 20 * 100) & "%"
    frmWork.prbPer.Value = intCount / 20 * 100
    DoEvents
Next intCount
For intCount = 16 To 20
    frmWork.lblPer = Str(intCount / 20 * 100) & "%"
    frmWork.prbPer.Value = intCount / 20 * 100
    DoEvents
    Map.AddFurniture intCount, 2, 1, "Furniture"
Next intCount
Print imlFurn.ListImages.Count
Unload frmWork
End Sub

Private Sub tmrWor_Timer()
    frmWork.lblPer = Str(intCount / 20 * 100) & "%"
    frmWork.prbPer.Value = intCount / 20 * 100
    DoEvents
End Sub

