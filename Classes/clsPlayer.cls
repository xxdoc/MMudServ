VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mNew As Boolean

Private mID As Integer
Private mIndex As Integer
Private mRoom As Long
Private mArea As Long

Private mName As String
Private mState As Integer
Private mPos As Integer

Private mLevel As Integer

Private mStam As Integer
Private mStr As Integer
Private mCoord As Integer
Private mIntel As Integer
Private mPass As String

Public Body As clsPlrBody

Public InRoom As clsRooms

Public Enum PLR_CONSTANTS
    PLR_NEW = 69                'Key value for new player creation
    PLR_CONNECTED = 1           'Key value for unverified connected players
    PLR_NAME = 2                'Key value for name entry recieved - unverified
    PLR_NAME_VERIFIED = 3       'Key value for name entry recieved - verified
    PLR_PASS = 4                'Key value for password receieved - unverified
    PLR_PASS_VERIFIED = 5       'Key value for password receieved - verified
    PLR_LOGGED_IN = 6           'Key value for player has logged in
End Enum

Public Enum PLR_POSITIONS
    PLR_STANDING = 1
    PLR_SITTING = 2
    PLR_RESTING = 3
    PLR_SLEEPING = 4
    PLR_AFK = 30
    PLR_DEAD = 666
    PLR_CORPSE = 999
End Enum

Public Property Get PlrNew() As Boolean
    PlrNew = mNew
End Property
Public Property Let PlrNew(ByVal vNewValue As Boolean)
    mNew = vNewValue
End Property

Public Property Get ID() As Integer
    ID = mID
End Property
Public Property Let ID(ByVal vNewValue As Integer)
    mID = vNewValue
End Property

Public Property Get PlrIndex() As Integer
    PlrIndex = mIndex
End Property
Public Property Let PlrIndex(ByVal vNewValue As Integer)
    mIndex = vNewValue
End Property

Public Property Get Room() As Long
    Room = mRoom
End Property
Public Property Let Room(ByVal vNewValue As Long)
    mRoom = vNewValue
End Property

Public Property Get Area() As Long
    Area = mArea
End Property
Public Property Let Area(ByVal vNewValue As Long)
    mArea = vNewValue
End Property

Public Property Get plrName() As String
    plrName = mName
End Property
Public Property Let plrName(ByVal vNewValue As String)
    mName = vNewValue
End Property

Public Property Get PlrState() As Integer
    PlrState = mState
End Property
Public Property Let PlrState(ByVal vNewValue As Integer)
    mState = vNewValue
End Property

Public Property Get Pos() As Integer
    Pos = mPos
End Property
Public Property Let Pos(ByVal vNewValue As Integer)
    mPos = vNewValue
End Property

'===============================================================================
'
Public Property Get Level() As Integer
    Level = mLevel
End Property
Public Property Let Level(ByVal vNewValue As Integer)
    mLevel = vNewValue
End Property



Public Property Get Coord() As Integer
    Coord = mCoord
End Property
Public Property Let Coord(ByVal vNewValue As Integer)
    mCoord = vNewValue
End Property

Public Property Get Intel() As Integer
    Intel = mIntel
End Property
Public Property Let Intel(ByVal vNewValue As Integer)
    mIntel = vNewValue
End Property

Public Property Get Strength() As Integer
    Strength = mStr
End Property
Public Property Let Strength(ByVal vNewValue As Integer)
    mStr = vNewValue
End Property

Public Property Get Stamina() As Integer
    Stamina = mStam
End Property
Public Property Let Stamina(ByVal vNewValue As Integer)
    mStam = vNewValue
End Property


Public Property Get Pass() As String
    Pass = mPass
End Property
Public Property Let Pass(ByVal vNewValue As String)
    mPass = vNewValue
End Property


Private Sub Class_Initialize()
    Set Body = New clsPlrBody
    Set InRoom = New clsRooms
End Sub
