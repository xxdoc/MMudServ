Attribute VB_Name = "HUB"
'Public NPCs() As clsNPCs
Public Mobiles() As clsMobiles
Public Commands() As clsCommands
Public Players() As clsPlayer
Public Screens() As clsScreens
'Public Settings As clsSettings
Public MudSys As clsSystem
'Public Stats As clsStats
Public LOG_COMMANDS As Boolean
Public Rooms() As clsRooms

Public MyConn As ADODB.Connection
Public MyRS As ADODB.Recordset
Public MyRS2 As ADODB.Recordset


Public Areas As clsAreas
Public ThePlayers As clsPlayers
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
Public Sub ParseText(Index As Integer, DataText As String)
    Dim arg1, arg2, arg3, arg4 As String
    Dim cAlph As String
    Dim plrName As String
    Dim plrPass As String
    
    Dim i As Integer
    
    Dim PlayerRS As ADODB.Recordset
    Set PlayerRS = New ADODB.Recordset
    RTrim$ (DataText)
    
    '=======================================================================================
    ' ToDo!
    '=======================================================================================
    ' Take this login code and break it up/move it.
    '===============================================
    If Not IS_LOGGED_IN(Index) Then

        If UCase$(DataText) = "NEW" Then
            HUB.Players(Index).PlrNew = True
            HUB.Players(Index).PlrState = PLR_NEW
            SendToUser Index, "New player created, please choose a first and last name for yourself." & vbCrLf
            Exit Sub
        
        ElseIf HUB.Players(Index).PlrState = PLR_CONNECTED Then
            arg1 = GetArg(DataText)
            If DataText = "" Then
                SendToUser Index, "%BYou must enter your full name.%n" & vbCrLf
                Exit Sub
            End If
            With PlayerRS
                DataText = Trim$(DataText)
                .open "SELECT * FROM PLR_DATA WHERE PLR_NAME='" & arg1 & " " & DataText & "'", MyConn, adOpenStatic, adLockOptimistic, adCmdText
                If .RecordCount >= 1 Then
                    SendToUser Index, "%bLoading your character...%n" & vbCrLf
                    
                    Dim ThePlayer As clsPlayer
                    Set ThePlayer = New clsPlayer
                    
                    ThePlayer.Area = !PLR_AREA
                    'ThePlayer.Body = !TODO!
                    ThePlayer.Coord = !PLR_COORD
                    ThePlayer.Stamina = !PLR_STAM
                    ThePlayer.Strength = !PLR_STR
                    
                    ThePlayer.ID = !PLR_ID
                    ThePlayer.Room = !PLR_ROOM
                    ThePlayer.Intel = !PLR_INTEL
                    ThePlayer.Level = !PLR_LEVEL
                    ThePlayer.Pass = !PLR_PASSWORD
                    ThePlayer.PlrIndex = Index
                    ThePlayer.Pos = !PLR_POS
                    ThePlayer.plrName = !PLR_NAME
                
                    ThePlayers.Add ThePlayer, Index
                                        
                    HUB.Players(Index).Coord = !PLR_COORD
                    
                    
                    HUB.Players(Index).ID = !PLR_ID
                    HUB.Players(Index).Intel = !PLR_INTEL
                    HUB.Players(Index).Level = !PLR_LEVEL
                    HUB.Players(Index).Pass = !PLR_PASSWORD
                    HUB.Players(Index).PlrIndex = Index
                    HUB.Players(Index).plrName = !PLR_NAME
                    HUB.Players(Index).Pos = !PLR_POS
                    HUB.Players(Index).Room = !PLR_ROOM
                    HUB.Players(Index).Area = !PLR_AREA
                                        
                    HUB.Players(Index).Stamina = !PLR_STAM
                    HUB.Players(Index).Strength = !PLR_STR
                    
                    HUB.Players(Index).PlrState = PLR_NAME_VERIFIED
                    SendToUser Index, "%GPlease enter your password to continue."
                End If
                .Close
            End With
            Exit Sub
        
        ElseIf HUB.Players(Index).PlrState = PLR_NEW Or HUB.Players(Index).PlrNew = True Then
        
            Exit Sub
        
        ElseIf HUB.Players(Index).PlrState = PLR_NAME_VERIFIED Then
            If HUB.Players(Index).Pass = DataText Then
                HUB.Players(Index).PlrState = PLR_LOGGED_IN
            
                HUB.ShowRoom Index
            End If
            Exit Sub
        End If
        
        HUB.Players(Index).PlrState = PLR_LOGGED_IN
        Exit Sub
    End If
    '///////////////////////////////////////////////////////////////////////////////////////
    
    
    If Not DataText = vbNullString Then 'And InStr(1, DataText, " ", vbTextCompare) Then
        arg1 = UCase$(GetArg(DataText))
    End If
    alph = mID$(arg1, 1, 1)
    
    If LOG_COMMANDS = True Then
        AddToLog "Command: " & arg1 & " processed - " & Date & " " & Time
    End If
    
    SkillTable Index, arg1, DataText
    
    
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
Public Function GetArg(ByRef DataText As String) As String
Dim Delim As Integer
    Delim = InStr(1, DataText, " ", vbTextCompare)
    If Delim = 0 Then
        GetArg = DataText
        DataText = vbNullString
    Else
        GetArg = mID(DataText, 1, Delim - 1)
        DataText = mID(DataText, Delim + 1)
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
Public Sub AddToLog(ByVal DataText As String)

    frmLog.List1.AddItem DataText
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
Public Sub ShowRoom(ByVal vIndex As Integer)
    Dim pRoom As Long
    Dim pArea As Long
    'With HUB.Rooms(HUB.Players(vIndex).Room)
    '    SendToUser vIndex, y1 & .Name & fRes1
    '    SendToUser vIndex, g1 & .LongDesc & fRes1
    'End With
    
    pRoom = ThePlayers.Item(vIndex).Room
    pArea = ThePlayers.Item(vIndex).Area
    
    SendToUser vIndex, y1 & Areas.Item(pArea).Rooms.Item(pRoom).RName & fRes1
    SendToUser vIndex, g1 & Areas.Item(pArea).Rooms.Item(pRoom).LongDesc & fRes1
    
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
'Public Sub SetPlayerRoom(ByVal vIndex As Integer, ByVal vNum As Integer)
'
'
'End Sub

Public Sub ShowHelpToUser(ByVal vIndex As Integer, ByVal vHelpItem As clsHelpItem)
Dim MyHelp As New clsHelpItem
Dim TheText As String
    Set MyHelp = vHelpItem
    
    If MyHelp.FROMFILE = False Then
        SendToUser vIndex, cl
        SendToUser vIndex, MyHelp.HELP_SCREEN
    Else
        SendToUser vIndex, cl
        Open App.Path & "\data\screens\help\" & MyHelp.HFileName For Input As #1
            Do While Not EOF(1)
                Input #1, TheText
                SendToUser vIndex, TheText
            Loop
        Close #1
    End If
    
End Sub
