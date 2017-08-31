VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================================
'          \\||//           | MWS | Concept \ Brian Knust
'           o o   - SEE YOU | AAE |  Coding  \  A.K.A.
'----o00o---(_)---o00o------| FRR |   Design  \  Godfather "Bizzle Knizzle"
'XXXXXXXXXXXXXXXXXXXXXXXXXXX| I V |---------------------------------------------------------
'XXXXXXXXXXXXXXXXXXXXXXXXXXX| A E | Copyright (c) 2012 BKPCS
'XXXXXXXXXXXXXXXXXXXXXXXXXXX|   R | All Rights Reserved
'===========================================================================================
' Notes: (New Notes on Top) FOLLOW THE FORMAT!!
' MM-DD-YYYY - (FIRST LAST) - <email@domain>
'=======================================================
' (ENTER NOTE TEXT HERE)
'===========================================================================================
' 05-31-2012 - (Brian Knust) - <BrianKnust@gmail.com>
'=======================================================
' TITLE:    THIS NEW CLASS
' NOTETEXT: This is the new exit class, it will be implemented as a collection in classExits
'
'===========================================================================================
'
Public Enum DIR_MOVEMENT
    DIR_N = 1   '1
    DIR_E       '2
    DIR_S       '3
    DIR_W       '4
    DIR_NE      '5
    DIR_SE      '6
    DIR_SW      '7
    DIR_NW      '8
    DIR_ClIMB   '9
    DIR_FLY     '10
    DIR_SWIM    '11
    DIR_DOWN    '12
    DIR_UP      '13
    DIR_SPEC    '14
End Enum

Private mAreaFrom As Long
Private mAreaTo As Long
Private mRoomFrom As Long
Private mRoomTo As Long
Private mDirect As Integer
Private mKeyWord As String

Private mDoor As Boolean
Private mClimb As Boolean
Private mFly As Boolean
Private mSwim As Boolean

Public Property Get AreaFrom() As Long
    AreaFrom = mAreaFrom
End Property
Public Property Let AreaFrom(ByVal vNewValue As Long)
    mAreaFrom = vNewValue
End Property

Public Property Get AreaTo() As Long
    AreaTo = mAreaTo
End Property
Public Property Let AreaTo(ByVal vNewValue As Long)
    mAreaTo = vNewValue
End Property


Public Property Get KeyWord() As String
    KeyWord = mKeyWord
End Property
Public Property Let KeyWord(ByVal vNewValue As String)
    mKeyWord = vNewValue
End Property


'===========================================================================================
' Property:
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date:
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Property Get RoomFrom() As Long
    RoomFrom = mRoomFrom
End Property
Public Property Let RoomFrom(ByVal vNewValue As Long)
    mRoomFrom = vNewValue
End Property


'===========================================================================================
' Property:
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date:
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Property Get RoomTo() As Long
    RoomTo = mRoomTo
End Property
Public Property Let RoomTo(ByVal vNewValue As Long)
    mRoomTo = vNewValue
End Property


'===========================================================================================
' Property:
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date:
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Property Get Direction() As Integer
    Direction = mDirect
End Property
Public Property Let Direction(ByVal vNewValue As Integer)
    mDirect = vNewValue
End Property

'===========================================================================================
' Property:
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date:
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Property Get Door() As Boolean
    Door = mDoor
End Property
Public Property Let Door(ByVal vNewValue As Boolean)
    mDoor = vNewValue
End Property


'===========================================================================================
' Property:
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date:
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Property Get Swim() As Boolean
    Swim = mSwim
End Property
Public Property Let Swim(ByVal vNewValue As Boolean)
    mSwim = vNewValue
End Property


'===========================================================================================
' Property:
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date:
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Property Get Climb() As Boolean
    Climb = mClimb
End Property
Public Property Let Climb(ByVal vNewValue As Boolean)
    mClimb = vNewValue
End Property


'===========================================================================================
' Property:
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date:
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Property Get Fly() As Boolean
    Fly = mFly
End Property
Public Property Let Fly(ByVal vNewValue As Boolean)
    mFly = vNewValue
End Property
