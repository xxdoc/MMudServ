Attribute VB_Name = "SendFuns"
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
Public Sub SendToUser(Index As Integer, DataText As String)
    If frmParent.Winsock1(Index).State = sckConnected Then
        frmParent.Winsock1(Index).SendData ParseANSI(DataText) & vbCrLf
    End If
End Sub


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
Public Sub Broadcast(Index As Integer, DataText As String)
Dim i As Integer
    For i = 1 To frmParent.Winsock1.UBound
        If frmParent.Winsock1(i).State = sckConnected Then
            frmParent.Winsock1(i).SendData vbCrLf & "Global: " & DataText & vbCrLf & vbCrLf
        End If
    Next i
End Sub


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
Public Sub SendToRoom(Index As Integer, DataText As String)
Dim i As Integer
    For i = 1 To frmParent.Winsock1.UBound
        If frmParent.Winsock1(i).State = sckConnected Then
            If HUB.Players(i).Room = HUB.Players(Index).Room Then
                If i = Index Then
                    SendToUser Index, DataText
                Else
                    SendToUser i, DataText
                End If
            End If
        End If
    Next i
End Sub
