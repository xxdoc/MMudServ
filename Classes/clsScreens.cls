VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mScreenData As String
Private mScreenName As String
Private mScreenID As Integer
Private mScreenLevel As Integer
Private mIS_File As Boolean
Private mSFileName As String

Public Property Get ScreenData() As String
    ScreenData = mScreenData
End Property
Public Property Let ScreenData(ByVal vNewValue As String)
    mScreenData = vNewValue
End Property

Public Property Get ScreenName() As String
    ScreenName = mScreenName
End Property
Public Property Let ScreenName(ByVal vNewValue As String)
    mScreenName = vNewValue
End Property

Public Property Get ScreenID() As Integer
    ScreenID = mScreenID
End Property
Public Property Let ScreenID(ByVal vNewValue As Integer)
    mScreenID = vNewValue
End Property

Public Property Get IS_File() As Boolean
    IS_File = mIS_File
End Property
Public Property Let IS_File(ByVal vNewValue As Boolean)
    mIS_File = vNewValue
End Property

Public Property Get SFileName() As String
    SFileName = mSFileName
End Property
Public Property Let SFileName(ByVal vNewValue As String)
    mSFileName = vNewValue
End Property

Public Property Get ScreenLevel() As Integer
    ScreenLevel = mScreenLevel
End Property
Public Property Let ScreenLevel(ByVal vNewValue As Integer)
    mScreenLevel = vNewValue
End Property
