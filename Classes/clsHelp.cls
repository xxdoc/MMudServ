VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHelps As New clsHelpItems

Public Property Set Helps(ByVal vNewValue As clsHelpItems)
    Set mHelps = vNewValue
End Property
Public Property Get Helps() As clsHelpItems
    Set Helps = mHelps
End Property
