Attribute VB_Name = "IS_FUNs"
Option Explicit

'===========================================================================================
' Sub:
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
Public Function IS_LOGGED_IN(ByVal vIndex As Integer) As Boolean
'ToDo:
    If HUB.Players(vIndex).PlrState = PLR_LOGGED_IN Then
        IS_LOGGED_IN = True
    Else
        IS_LOGGED_IN = False
    End If
End Function

'===========================================================================================
' Sub:
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
Public Function IS_AFK(ByVal vIndex As Integer) As Boolean
'ToDo:
End Function


'===========================================================================================
' Sub:
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
Public Function IS_NEW(ByVal vIndex As Integer) As Boolean
    If HUB.Players(vIndex).PlrState = PLR_NEW Then
        IS_NEW = True
    Else
        IS_NEW = False
    End If
End Function
