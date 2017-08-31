VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String
Private mPos As Integer
Private mCode As String
Private mLevel As Integer
Private mLog As Boolean

Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(ByVal vNewValue As String)
    mName = vNewValue
End Property

Public Property Get Code() As String
    Code = mCode
End Property
Public Property Let Code(ByVal vNewValue As String)
    mCode = vNewValue
End Property

Public Property Get Level() As Integer
    Level = mLevel
End Property
Public Property Let Level(ByVal vNewValue As Integer)
    mLevel = vNewValue
End Property

Public Property Get Pos() As Integer
    Pos = mPos
End Property
Public Property Let Pos(ByVal vNewValue As Integer)
    mPos = vNewValue
End Property

Public Property Get Log() As Boolean
    Log = mLog
End Property
Public Property Let Log(ByVal vNewValue As Boolean)
    mLog = vNewValue
End Property
