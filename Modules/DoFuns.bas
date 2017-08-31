Attribute VB_Name = "DoFuns"
Option Explicit

'===========================================================================================
' Sub:      DoNorth (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-21-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to handle movement of a character or wandering NPC from one room to
'           another room that lies to the north
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoNorth(Index As Integer, DataText As String)
    Dim i As Integer
    Dim VNum As Integer
    
    'For i = 1 To UBound(HUB.Rooms)
    '    If HUB.Rooms(i).VNum = HUB.Players(Index).Room Then Exit For
    'Next i
    
    'If HUB.Rooms(i).GetExit(DIR_NORTH, VNum) Then
    '    SendToUser Index, "%BYou go north" & vbCrLf & vbCrLf
    '    HUB.Players(Index).Room = VNum
    '    HUB.ShowRoom Index
    'End If
    
    NewDoNorth Index, DataText

End Sub

Public Sub NewDoNorth(Index As Integer, DataText As String)
    'ToDo:
    Dim ExCount As Integer
    Dim i As Integer
    Dim ExIndex As Long
    
    ExCount = Areas.Item(ThePlayers.Item(Index).Area).Rooms.Item(ThePlayers.Item(Index).Room).Exits.Count
    
    If ExCount > 0 Then
        For i = 1 To ExCount
            If Areas.Item(ThePlayers.Item(Index).Area).Rooms.Item(ThePlayers.Item(Index).Room).Exits.Item(i).Direction = DIR_NORTH Then
                ThePlayers.Item(Index).Area = Areas.Item(ThePlayers.Item(Index).Area).Rooms.Item(ThePlayers.Item(Index).Room).Exits.Item(i).AreaTo
                ThePlayers.Item(Index).Room = Areas.Item(ThePlayers.Item(Index).Area).Rooms.Item(ThePlayers.Item(Index).Room).Exits.Item(i).RoomTo
                ShowRoom Index
            Else
                SendToUser Index, "%rThere is no exit in that direction!" & vbCrLf & vbCrLf
            End If
            Debug.Print "Exit Direction: " & Areas.Item(ThePlayers.Item(Index).Area).Rooms.Item(ThePlayers.Item(Index).Room).Exits.Item(i).Direction
        Next i
    End If
    
End Sub


'===========================================================================================
' Sub:      DoSouth (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-21-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to handle movement of a character or wandering NPC from one room to
'           another room that lies to the south
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoSouth(Index As Integer, DataText As String)
    Dim i As Integer
    Dim VNum As Integer
    
    For i = 1 To UBound(HUB.Rooms)
        If HUB.Rooms(i).VNum = HUB.Players(Index).Room Then Exit For
    Next i
    
    If HUB.Rooms(i).GetExit(DIR_SOUTH, VNum) Then
        SendToUser Index, "%BYou go south" & vbCrLf & vbCrLf
        HUB.Players(Index).Room = VNum
        HUB.ShowRoom Index
    End If
End Sub


'===========================================================================================
' Sub:      DoWest (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-21-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to handle movement of a character or wandering NPC from one room to
'           another room that lies to the west
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoWest(Index As Integer, DataText As String)
    Dim i As Integer
    Dim VNum As Integer
    
    For i = 1 To UBound(HUB.Rooms)
        If HUB.Rooms(i).VNum = HUB.Players(Index).Room Then Exit For
    Next i
    
    If HUB.Rooms(i).GetExit(DIR_WEST, VNum) Then
        SendToUser Index, "%BYou go west" & vbCrLf & vbCrLf
        HUB.Players(Index).Room = VNum
        HUB.ShowRoom Index
    End If
End Sub


'===========================================================================================
' Sub:      DoEast (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-21-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to handle movement of a character or wandering NPC from one room to
'           another room that lies to the east
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoEast(Index As Integer, DataText As String)
    Dim i As Integer
    Dim VNum As Integer
    
    For i = 1 To UBound(HUB.Rooms)
        If HUB.Rooms(i).VNum = HUB.Players(Index).Room Then Exit For
    Next i
    
    If HUB.Rooms(i).GetExit(DIR_EAST, VNum) Then
        SendToUser Index, "%BYou go east" & vbCrLf & vbCrLf
        HUB.Players(Index).Room = VNum
        HUB.ShowRoom Index
    End If
End Sub


'===========================================================================================
' Sub:      DoHelp (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-21-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to determine which help file is to be shown to a player requesting
'           help and show the appropriate screen.
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoHelp(Index As Integer, DataText As String)
Dim i As Integer
Dim HelpItem As New clsHelpItem
Dim HelpTerm, args() As String

    
    SendToUser Index, "%BYou ask for help on " & DataText & vbCrLf & vbCrLf
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        ReDim Preserve args(UBound(args) + 1)
        i = i + 1
        args(i) = GetArg(DataText)
    Loop
    
    HelpTerm = vbNullString
    
    For i = 1 To UBound(args)
        HelpTerm = HelpTerm & args(i)
        If Not (i = UBound(args)) Then HelpTerm = HelpTerm & " "
    Next i
    
    If (frmParent.TheHelp.Helps.Item(HelpItem, HelpTerm)) = False Then
        SendToUser Index, "%RHelp on " & HelpTerm & " was not found." & vbCrLf & vbCrLf
    Else
        With HelpItem
            HUB.ShowHelpToUser Index, HelpItem
        End With
    End If
    
End Sub

'===========================================================================================
' Sub:      DoNorthEast (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-21-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to handle movement of a character or wandering NPC from one room to
'           another that lies to the North-east.
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoNorthEast(Index As Integer, DataText As String)
    Dim i As Integer
    Dim VNum As Integer
    
    For i = 1 To UBound(HUB.Rooms)
        If HUB.Rooms(i).VNum = HUB.Players(Index).Room Then Exit For
    Next i
    
    If HUB.Rooms(i).GetExit(DIR_NORTHEAST, VNum) Then
        SendToUser Index, "%BYou go northeast" & vbCrLf & vbCrLf
        HUB.Players(Index).Room = VNum
        HUB.ShowRoom Index
    End If
End Sub


'===========================================================================================
' Sub:      DoNorthWest (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-21-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to handle movement of a character or wandering NPC from one room to
'           another that lies to the North-west
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoNorthWest(Index As Integer, DataText As String)
    Dim i As Integer
    Dim VNum As Integer
    
    For i = 1 To UBound(HUB.Rooms)
        If HUB.Rooms(i).VNum = HUB.Players(Index).Room Then Exit For
    Next i
    
    If HUB.Rooms(i).GetExit(DIR_NORTHWEST, VNum) Then
        SendToUser Index, "%BYou go northwest" & vbCrLf & vbCrLf
        HUB.Players(Index).Room = VNum
        HUB.ShowRoom Index
    End If
End Sub


'===========================================================================================
' Sub:      DoSouthEast (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-21-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to handle movement of a character or wandering NPC from one room to
'           another that lies to the South-east
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoSouthEast(Index As Integer, DataText As String)
    Dim i As Integer
    Dim VNum As Integer
    
    For i = 1 To UBound(HUB.Rooms)
        If HUB.Rooms(i).VNum = HUB.Players(Index).Room Then Exit For
    Next i
    
    If HUB.Rooms(i).GetExit(DIR_SOUTHEAST, VNum) Then
        SendToUser Index, "%BYou go southeast" & vbCrLf & vbCrLf
        HUB.Players(Index).Room = VNum
        HUB.ShowRoom Index
    End If
End Sub


'===========================================================================================
' Sub:      DoSouthWest (Integer, String)x
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Sub-routine to handle movement of a character or wandering NPC from one room to
'           another that lies to the South-west
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoSouthWest(Index As Integer, DataText As String)
    Dim i As Integer
    Dim VNum As Integer
    
    For i = 1 To UBound(HUB.Rooms)
        If HUB.Rooms(i).VNum = HUB.Players(Index).Room Then Exit For
    Next i
    
    If HUB.Rooms(i).GetExit(DIR_SOUTHWEST, VNum) Then
        SendToUser Index, "%BYou go southwest" & vbCrLf & vbCrLf
        HUB.Players(Index).Room = VNum
        HUB.ShowRoom Index
    End If
End Sub


'===========================================================================================
' Sub:      DoSay (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handle telling all in a room what a player has said.
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoSay(Index As Integer, DataText As String)
    SendToUser Index, "You say: '" & DataText & "'" & vbCrLf
End Sub


'===========================================================================================
' Sub:      DoSell (Integer, String)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handle a player selling an item to an NPC shop.
'
'===========================================================================================
' Notes:    (Currently not implemented)
'
'===========================================================================================
'
Public Sub DoSell(Index As Integer, DataText As String)
'ToDo:
Dim arg1, arg2 As String

End Sub


'===========================================================================================
' Sub:      DoScore ()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handle the score feature (Not yet implemented)
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub DoScore(Index As Integer, DataText As String)
'ToDo:
Dim arg1 As String
Dim sScr As String

    SendToUser Index, "Score Feature Not Yet Available" & vbCrLf
    

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
Public Sub DoShoot(Index As Integer, DataText As String)
Dim arg1 As String

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
Public Sub DoShout(Index As Integer, DataText As String)
Dim i As Integer
    For i = 1 To frmParent.Winsock1.UBound
        If i = Index Then
            SendToUser Index, "%rYou shout: '" & DataText & "'"
        Else
            SendToUser i, "%r" & HUB.Players(Index).plrName & " shouts: '" & DataText & "'"
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
Public Sub DoYell(Index As Integer, DataText As String)
Dim i As Integer
    For i = 1 To frmParent.Winsock1.UBound
        If i = Index Then
            SendToUser Index, "%RYou yell: '" & DataText & "'%n"
        Else
            SendToUser i, "%R[name] yells: '" & DataText & "'%n"
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
Public Sub DoHire(Index As Integer, DataText As String)
'ToDo:
Dim arg1, arg2 As String

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
Public Sub DoScreen(Index As Integer, DataText As String)
    ShowScreenToUser Index, DataText
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
Public Sub DoReload(Index As Integer, DataText As String)
Dim arg1 As String

    arg1 = GetArg(DataText)

    If UCase$(arg1) = "SCREENS" Then
        frmParent.LoadScreens
        SendToUser Index, "%bScreens re-loaded%n"
        AddToLog "Screens reloaded by " & HUB.Players(Index).plrName & " - " & Date & Time
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
Public Sub DoCreate(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoMCreate(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoOCreate(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoRCreate(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoHCreate(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoCoCreate(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoMEdit(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoOEdit(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoREdit(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoHEdit(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoCoEdit(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoMELock(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoOELock(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoRELock(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoHELock(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
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
Public Sub DoCOELock(Index As Integer, DataText As String)
Dim args() As String
Dim i As Integer
    
    i = 0
    ReDim Preserve args(0)
    Do While Not (DataText = "")
        args(i) = GetArg(DataText)
        i = i + 1
        If Not (DataText = "") Then ReDim Preserve args(UBound(args) + 1)
    Loop
    
    'ToDo:
End Sub

