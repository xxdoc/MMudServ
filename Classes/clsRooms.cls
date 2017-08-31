VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Exits() As clsExits

Private mRMIndex As Integer
Private mVNum As Integer
Private mName As String
Private mShort As String
Private mLong As String

Public Property Get RoomIndex() As Integer
    RoomIndex = mRMIndex
End Property
Public Property Let RoomIndex(ByVal vNewValue As Integer)
    mRMIndex = vNewValue
End Property


Public Property Get vNum() As Integer
    vNum = mVNum
End Property
Public Property Let vNum(ByVal vNewValue As Integer)
    mVNum = vNewValue
End Property


Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(ByVal vNewValue As String)
    mName = vNewValue
End Property

Public Property Get ShortDesc() As String
    ShortDesc = mShort
End Property
Public Property Let ShortDesc(ByVal vNewValue As String)
    mShort = vNewValue
End Property


Public Property Get LongDesc() As String
    LongDesc = mLong
End Property
Public Property Let LongDesc(ByVal vNewValue As String)
    mLong = vNewValue
End Property

Public Sub AddExit(ByVal vRoomTo As Integer, ByVal vDirection As Integer, ByVal vDoor As Boolean, Optional ByVal vClimb As Boolean, Optional ByVal vFly As Boolean)
    ReDim Preserve Exits(UBound(Exits) + 1)
    Set Exits(UBound(Exits)) = New clsExits
    
    With Exits(UBound(Exits))
        .RoomTo = vRoomTo
        .RoomFrom = Me.vNum
        .Direction = vDirection
        .Door = vDoor
        .Climb = vClimb
        .Fly = vFly
    End With
    
End Sub

Public Function GetExit(ByVal vDirection As Integer, ByRef vNum As Integer) As Boolean
Dim i As Integer
    For i = 1 To UBound(Exits)
        If Exits(i).Direction = vDirection Then
            GetExit = True
            vNum = Exits(i).RoomTo
            Exit For
        Else
            GetExit = False
        End If
    Next i
End Function

Public Function GetExit2(ByVal vDirection, ByRef vExit As clsExits) As Boolean
Dim i As Integer
    For i = 1 To UBound(Exits)
        If Exits(i).Direction = vDirection Then
                   GetExit2 = True
        End If
    Next i
End Function

Public Sub RemoveExit()
    
End Sub

Private Sub Class_Initialize()
    ReDim Preserve Exits(1)
    Set Exits(1) = New clsExits
End Sub
