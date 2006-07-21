Attribute VB_Name = "MainOS"
Option Explicit
Public Mons As New GMotion.Motion
Public Map As New GMotion.Map
Public Load As New GMotion.Load
Public Member As New GMotion.Member
Public SEvents As New SEvents
Public intFps As Integer          ' Fps
Public Company As String          'имя компании
Public Level As Integer           'номер уровня
Public NumPersonag As Integer     'кол. персонажей
Public Files As Byte           'кол. папок
Public Xp As Integer, Yp As Integer  ' кооординаты местности
Public Lifes As Byte                  'жизни
Public Money As Integer              'кол. денег
Public startLifes As Byte                 'на начало уровня жизни
Public startMoney As Integer              'на начало уровня кол. денег
Public SaveMode As Boolean           'вид страницы загрузки

Sub Main()
    Member.MCatch = False
    intFps = 24
    frmMain.Show
End Sub
