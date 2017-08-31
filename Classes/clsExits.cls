VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum DIR_MOVEMENTS
    DIR_NORTH = 1
    DIR_EAST
    DIR_SOUTH
    DIR_WEST
    DIR_NORTHEAST
    DIR_SOUTHEAST
    DIR_SOUTHWEST
    DIR_NORTHWEST
End Enum


Private mRoomFrom As Integer
Private mRoomTo As Integer
Private mDirect As Integer

Private mDoor As Boolean
Private mClimb As Boolean
Private mFly As Boolean


Public Property Get Direction() As Integer
    Direction = mDirect
End Property
Public Property Let Direction(ByVal vNewValue As Integer)
    mDirect = vNewValue
End Property


Public Property Get RoomFrom() As Integer
    RoomFrom = mRoomFrom
End Property
Public Property Let RoomFrom(ByVal vNewValue As Integer)
    mRoomFrom = vNewValue
End Property


Public Property Get RoomTo() As Integer
    RoomTo = mRoomTo
End Property
Public Property Let RoomTo(ByVal vNewValue As Integer)
    mRoomTo = vNewValue
End Property


Public Property Get Door() As Boolean
    Door = mDoor
End Property
Public Property Let Door(ByVal vNewValue As Boolean)
    mDoor = vNewValue
End Property


Public Property Get Climb() As Boolean
    Climb = mClimb
End Property
Public Property Let Climb(ByVal vNewValue As Boolean)
    mClimb = vNewValue
End Property


Public Property Get Fly() As Boolean
    Fly = mFly
End Property
Public Property Let Fly(ByVal vNewValue As Boolean)
    mFly = vNewValue
End Property
