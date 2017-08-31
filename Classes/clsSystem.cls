VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMaxLevel As Integer
Private mPkLoot As Boolean
Private mWaitForAuth As Boolean
Private mStunPVP As Integer
Private mStunReg As Integer
Private mDodgeMod As Integer
Private mParryMod As Integer
Private mDamPVP As Integer
Private mDamPVM As Integer
Private mDamMVP As Integer
Private mBuild As Integer
Private mProto As Integer
Private mMudName As String

Public Property Get Build() As Integer
    Build = mBuild
End Property
Public Property Let Build(ByVal vNewValue As Integer)
    mBuild = vNewValue
End Property

Public Property Get Proto() As Integer
    Proto = mProto
End Property
Public Property Let Proto(ByVal vNewValue As Integer)
    mProto = vNewValue
End Property

Public Property Get DamMVP() As Integer
    DamMVP = mDamMVP
End Property
Public Property Let DamMVP(ByVal vNewValue As Integer)
    mDamMVP = vNewValue
End Property

Public Property Get DamPVM() As Integer
    DamPVM = mDamPVM
End Property
Public Property Let DamPVM(ByVal vNewValue As Integer)
    mDamPVM = vNewValue
End Property

Public Property Get DamPVP() As Integer
    DamPVP = mDamPVP
End Property
Public Property Let DamPVP(ByVal vNewValue As Integer)
    mDamPVP = vNewValue
End Property

Public Property Get ParryMod() As Integer
    ParryMod = mParryMod
End Property
Public Property Let ParryMod(ByVal vNewValue As Integer)
    mParryMod = vNewValue
End Property

Public Property Get DodgeMod() As Integer
    DodgeMod = mDodgeMod
End Property
Public Property Let DodgeMod(ByVal vNewValue As Integer)
    mDodgeMod = vNewValue
End Property

Public Property Get StunReg() As Integer
    StunReg = mStunReg
End Property
Public Property Let StunReg(ByVal vNewValue As Integer)
    mStunReg = vNewValue
End Property

Public Property Get StunPVP() As Integer
    StunPVP = mStunPVP
End Property
Public Property Let StunPVP(ByVal vNewValue As Integer)
    mStunPVP = vNewValue
End Property

Public Property Get WaitForAuth() As Boolean
    WaitForAuth = mWaitForAuth
End Property
Public Property Let WaitForAuth(ByVal vNewValue As Boolean)
    mWaitForAuth = vNewValue
End Property

Public Property Get PKLoot() As Boolean
    PKLoot = mPkLoot
End Property
Public Property Let PKLoot(ByVal vNewValue As Boolean)
    mPkLoot = vNewValue
End Property

Public Property Get MudName() As String
    MudName = mMudName
End Property
Public Property Let MudName(ByVal vNewValue As String)
    mMudName = vNewValue
End Property

Public Property Get MaxLevel() As Integer
    MaxLevel = mMaxLevel
End Property
Public Property Let MaxLevel(ByVal vNewValue As Integer)
    mMaxLevel = vNewValue
End Property
