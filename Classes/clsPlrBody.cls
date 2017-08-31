VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlrBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHead As Integer
Private mNeck As Integer
Private mChest As Integer
Private mArms As Integer
Private mLegs As Integer
Private mFeet As Integer
Private mRHand As Integer
Private mLHand As Integer

Public Property Get Head() As Integer
    Head = mHead
End Property
Public Property Let Head(ByVal vNewValue As Integer)
    mHead = vNewValue
End Property

Public Property Get Neck() As Integer
    Neck = mNeck
End Property
Public Property Let Neck(ByVal vNewValue As Integer)
    mNeck = vNewValue
End Property

Public Property Get Chest() As Integer
    Chest = mChest
End Property
Public Property Let Chest(ByVal vNewValue As Integer)
    mChest = vNewValue
End Property

Public Property Get Arms() As Integer
    Arms = mArms
End Property
Public Property Let Arms(ByVal vNewValue As Integer)
    mArms = vNewValue
End Property

Public Property Get Legs() As Integer
    Legs = mLegs
End Property
Public Property Let Legs(ByVal vNewValue As Integer)
    mLegs = vNewValue
End Property

Public Property Get Feet() As Integer
    Feet = mFeet
End Property
Public Property Let Feet(ByVal vNewValue As Integer)
    mFeet = vNewValue
End Property

Public Property Get RHand() As Integer
    RHand = mRHand
End Property
Public Property Let RHand(ByVal vNewValue As Integer)
    mRHand = vNewValue
End Property

Public Property Get LHand() As Integer
    LHand = mLHand
End Property
Public Property Let LHand(ByVal vNewValue As Integer)
    mLHand = vNewValue
End Property
