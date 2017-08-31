VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAreas"
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
Private mCol As Collection


Public Function Add(ByRef aArea As clsArea, ByVal vKey As Variant) As clsArea
    'Dim objNew As clsArea
    'Set objNew = New clsArea
    
    'objNew.AName = aArea.AName
    'objNew.HVnum = aArea.HVnum
    'objNew.LVnum = aArea.LVnum
    'objNew.ResMessage = aArea.ResMessage
    'objNew.ResTime = aArea.ResTime
    
    mCol.Add aArea, "" & vKey & ""
    
    'Set Add = objNew
    'Set objNew = Nothing
End Function


Public Property Get Item(vIndexKey As Variant) As clsArea
    Set Item = mCol(vIndexKey)
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
