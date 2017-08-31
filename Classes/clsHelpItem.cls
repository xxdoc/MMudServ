VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHelpItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mHELPNAME As String
Private mHELP_SYNTAX As String
Private mHELP_SCREEN As String
Private mFROMFILE As Boolean
Private mHFileName As String


Public Property Get HELPNAME() As String
    HELPNAME = mHELPNAME
End Property
Public Property Let HELPNAME(ByVal vNewValue As String)
    mHELPNAME = vNewValue
End Property

Public Property Get HELP_SYNTAX() As String
    HELP_SYNTAX = mHELP_SYNTAX
End Property
Public Property Let HELP_SYNTAX(ByVal vNewValue As String)
    mHELP_SYNTAX = vNewValue
End Property

Public Property Get HELP_SCREEN() As String
    HELP_SCREEN = mHELP_SCREEN
End Property
Public Property Let HELP_SCREEN(ByVal vNewValue As String)
    mHELP_SCREEN = vNewValue
End Property

Public Property Get FROMFILE() As Boolean
    FROMFILE = mFROMFILE
End Property
Public Property Let FROMFILE(ByVal vNewValue As Boolean)
    mFROMFILE = vNewValue
End Property

Public Property Get HFileName() As String
    HFileName = mHFileName
End Property
Public Property Let HFileName(ByVal vNewValue As String)
    mHFileName = vNewValue
End Property
