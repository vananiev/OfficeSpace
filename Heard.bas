Attribute VB_Name = "Heard"
Option Explicit
Public Load As New Load
Public Mons As New Motion
Public Map As New Map
Public FerstLoad As New FerstLoad
Public X(15) As Single            ' текущая x корд. данного персонажа(пикселы)
Public Y(15) As Single            ' текущая y корд. данного персонажа(пикселы)
Public intPos(15) As Integer      ' текущая позиция картинки
Public intMaxPos(15) As Integer   ' мах позиция картинки
Public strDir(15) As String       ' направление движения
Public blnMove(15) As Boolean     ' выполнение движения в текущий момент
Public Personag As Integer        ' активный персонаж
Public Catch As Boolean           'флаг, признак того, что персонажа поймали
Public ServiseForm As Object      ' ссылка на рабочую форму
Public objReturn As Object        ' ссылка на нерабочую форму
Public SeeDistance(15) As Integer ' радиус зрения
Public Xp As Integer, Yp As Integer  ' кооординаты местности
Public Const SetMajor = 11           ' количество картинок на персонажа
Public Place(15, 42, 3) As Byte   ' информация об объектах (0-вид, 1-начальная X координата, 2-начальная Y координата,3-информация об плитке)
Public FurnW(223) As Byte        ' ширина интерьера
Public FurnH(223) As Byte        ' длинна интерьера
Public maxFurn As Integer            ' мах рисунок мебели
Public GameDir As String               ' папка игры
Public intCount As Integer

Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal X As Long, ByVal у As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long

Sub Main()

End Sub

