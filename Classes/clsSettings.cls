VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===========================================================================================
'
'===========================================================================================
'
Private mName As String
Private mMaxLevel As Integer


'===========================================================================================
'
Public Property Get MudName() As String
    MudName = mName
End Property
Public Property Let MudName(ByVal vNewvalue As String)
    mName = vNewvalue
End Property


'===========================================================================================
'
Public Property Get MaxLevel() As Integer

End Property
Public Property Let MaxLevel(ByVal vNewvalue As Integer)

End Property
