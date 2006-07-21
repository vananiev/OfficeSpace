VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Install"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Эта программа извлечет мультимедийные файлы, необходимые для работы прграммы, а затем установит Office Space на ваш компьютер."
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mfsysObject As New Scripting.FileSystemObject
Dim strDir As String
Dim strFileName As String

Private Sub cmdInstall_Click()
    Dim mFolder As Folder
    strDir = mfsysObject.GetSpecialFolder(WindowsFolder)
    strDir = Left(strDir, 3) ' отсекаем лишнее
    'Cоздание папок
    On Error Resume Next
    MkDir strDir & "Program Files\Softway\"
    On Error Resume Next
    MkDir strDir & "Program Files\Softway\OfficeSpace\"
    Set mFolder = mfsysObject.GetFolder(App.Path & "\Sound\")
    mFolder.Copy strDir
    Set mFolder = mfsysObject.GetFolder(App.Path & "\Image\")
    mFolder.Copy strDir & "\Program Files\Softway\OfficeSpace\"
    Set mFolder = mfsysObject.GetFolder(App.Path & "\Companies\")
    mFolder.Copy strDir & "\Program Files\Softway\OfficeSpace\"
    Set mFolder = mfsysObject.GetFolder(App.Path & "\Save Games\")
    mFolder.Copy strDir & "\Program Files\Softway\OfficeSpace\"
    On Error GoTo OneExe
    Shell App.Path & "\Setup.exe"
    Unload Main
    End
    Exit Sub
OneExe:
    FileCopy App.Path & "\OfficeSpace.exe", strDir & "\Program Files\Softway\OfficeSpace\OfficeSpace.exe"
    Unload Main
    End
End Sub
