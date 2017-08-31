VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoom"
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
Private mAIndex As Long
Private mRMIndex As Long
Private mVNum As Long
Private mName As String
Private mShort As String
Private mLong As String
Private mExits As classExits

Public Property Set Exits(ByVal vNewValue As classExits)
    Set mExits = vNewValue
End Property
Public Property Get Exits() As classExits
    Set Exits = mExits
End Property

'===========================================================================================
' Property: AIndex
' PropType: Long
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     5-25-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 5-25-2012
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Property Get AIndex() As Long
    AIndex = mAIndex
End Property
Public Property Let AIndex(ByVal vNewValue As Long)
    mAIndex = vNewValue
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
Public Property Get RMIndex() As Long
    RMIndex = mRMIndex
End Property
Public Property Let RMIndex(ByVal vNewValue As Long)
    mRMIndex = vNewValue
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
Public Property Get VNum() As Long
    VNum = mVNum
End Property
Public Property Let VNum(ByVal vNewValue As Long)
    mVNum = vNewValue
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
Public Property Get RName() As String
    RName = mName
End Property
Public Property Let RName(ByVal vNewValue As String)
    mName = vNewValue
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
Public Property Get ShortDesc() As String
    ShortDesc = mShort
End Property
Public Property Let ShortDesc(ByVal vNewValue As String)
    mShort = vNewValue
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
Public Property Get LongDesc() As String
    LongDesc = mLong
End Property
Public Property Let LongDesc(ByVal vNewValue As String)
    mLong = vNewValue
End Property

Private Sub Class_Initialize()
    Set mExits = New classExits
End Sub
Private Sub Class_Terminate()
    Set mExits = Nothing
End Sub
