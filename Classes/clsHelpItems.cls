VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHelpItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection


Public Function Add(HELPNAME As String, FROMFILE As Boolean, Optional HFileName As String, Optional HELP_SYNTAX As String, Optional HELP_SCREEN As String, Optional sKey As String) As clsHelpItem

    Dim objNew As clsHelpItem
    Set objNew = New clsHelpItem
    
    objNew.HELPNAME = HELPNAME
    objNew.FROMFILE = FROMFILE
    If FROMFILE = True Then
        objNew.HFileName = HFileName
        objNew.HELP_SYNTAX = vbNull
        objNew.HELP_SCREEN = vbNull
    Else
        objNew.HFileName = vbNull
        objNew.HELP_SYNTAX = HELP_SYNTAX
        objNew.HELP_SCREEN = HELP_SCREEN
    End If
    
    mCol.Add objNew, HELPNAME
    
    Set Add = objNew
    Set objNew = Nothing

End Function


'===========================================================================================
' Property: Get Item (ByRef clsHelpItem, Variant) As Boolean
' Returns:  Boolean
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript:
'
'===========================================================================================
' Notes:    Let's do something different here. Usually for a collection we need to return
'           an object from a collection. However, with this example we also need to know if
'           fetching the item was successful. The only error we expect to get is that the
'           object can not be found with the given key. By passing a reference (the memory
'           location) of the object we want to recieve the object in, and returning a
'           a boolean, we are able to "return" two values of different types, and only use
'           the data in the object if the "Get" returns 'True'.
'
'===========================================================================================
'
Public Property Get Item(ByRef hiHelpItem As clsHelpItem, vIndexKey As Variant) As Boolean
On Error GoTo Error_clsHelpItems_Get_Item:
    Set hiHelpItem = mCol(vIndexKey)
    Item = True
Exit Property
Error_clsHelpItems_Get_Item:
    Item = False
    
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vIndexKey As Variant)
    mCol.Remove vIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

