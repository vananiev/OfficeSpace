VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get MServiseForm() As Object
    Set MServiseForm = ServiseForm
End Property

Public Property Let MServiseForm(ByVal vNewValue As Object)
    Set ServiseForm = vNewValue
End Property

Public Property Get MCatch() As Boolean
    MCatch = Catch
End Property

Public Property Let MCatch(ByVal vNewValue As Boolean)
    Catch = vNewValue
End Property

Public Property Get mX(ByVal Index As Integer) As Single
    mX = X(Index)
End Property

Public Property Let mX(ByVal Index As Integer, ByVal vNewValue As Single)
    X(Index) = vNewValue
End Property

Public Property Get mY(ByVal Index As Integer) As Single
    mY = Y(Index)
End Property

Public Property Let mY(ByVal Index As Integer, ByVal vNewValue As Single)
    Y(Index) = vNewValue
End Property


Public Property Get Position(ByVal Index As Integer) As Integer
    Position = intPos(Index)
End Property

Public Property Let Position(ByVal Index As Integer, ByVal vNewValue As Integer)
    intPos(Index) = vNewValue
End Property

Public Property Get Direction(ByVal Index As Integer) As String
    Direction = strDir(Index)
End Property

Public Property Let Direction(ByVal Index As Integer, ByVal vNewValue As String)
    strDir(Index) = vNewValue
End Property

Public Property Get Move(ByVal Index As Integer) As Boolean
    Move = blnMove(Index)
End Property

Public Property Let Move(ByVal Index As Integer, ByVal vNewValue As Boolean)
    blnMove(Index) = vNewValue
End Property

Public Property Get MSeeDistance(ByVal Index As Integer) As Integer
    MSeeDistance = SeeDistance(Index)
End Property

Public Property Let MSeeDistance(ByVal Index As Integer, ByVal vNewValue As Integer)
    SeeDistance(Index) = vNewValue
End Property

Public Property Get MPlace(ByVal Xp As Integer, ByVal Yp As Integer, ByVal Dir As Integer) As Integer
    MPlace = Place(Xp, Yp, Dir)
End Property

Public Property Let MPlace(ByVal Xp As Integer, ByVal Yp As Integer, ByVal Dir As Integer, ByVal vNewValue As Integer)
    MPlace(Xp, Yp, Dir) = vNewValue
End Property

Public Property Get GameDirectory() As String
    GameDirectory = GameDir
End Property

Public Property Let GameDirectory(ByVal vNewValue As String)
    GameDir = vNewValue
End Property

Public Property Get MaxFurniture() As Integer
    MaxFurniture = maxFurn
End Property
