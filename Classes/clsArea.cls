VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArea"
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
'
'
Private mRooms As classRooms
Private mHVnum As Long
Private mLVnum As Long
Private mResTime As Integer
Private mResMessage As String
Private mAName As String

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
Public Property Get AName() As String
    AName = mAName
End Property
Public Property Let AName(ByVal vNewValue As String)
    mAName = vNewValue
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
Public Property Get ResMessage() As String
    ResMessage = mResMessage
End Property
Public Property Let ResMessage(ByVal vNewValue As String)
    mResMessage = vNewValue
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
Public Property Get ResTime() As Integer
    ResTime = mResTime
End Property
Public Property Let ResTime(ByVal vNewValue As Integer)
    mResTime = vNewValue
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
Public Property Get LVnum() As Long
    LVnum = mLVnum
End Property
Public Property Let LVnum(ByVal vNewValue As Long)
    mLVnum = vNewValue
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
Public Property Get HVnum() As Long
    HVnum = mHVnum
End Property
Public Property Let HVnum(ByVal vNewValue As Long)
    mHVnum = vNewValue
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
Public Property Set Rooms(ByVal vNewValue As classRooms)
    Set mRooms = vNewValue
End Property
Public Property Get Rooms() As classRooms
    Set Rooms = mRooms
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
Private Sub Class_Initialize()
    Set mRooms = New classRooms
End Sub
Private Sub Class_Terminate()
    Set mRooms = Nothing
End Sub
