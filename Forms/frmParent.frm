VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmParent 
   BackColor       =   &H8000000C&
   Caption         =   "Mafia: World Server"
   ClientHeight    =   8205
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9840
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox sbMud 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   9780
      TabIndex        =   0
      Top             =   7965
      Width           =   9840
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4040
   End
   Begin VB.Menu mnuFileMain 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuFileStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSystemMain 
      Caption         =   "&System"
      Begin VB.Menu mnuSysCommands 
         Caption         =   "&Commands"
      End
      Begin VB.Menu mnuScreensMain 
         Caption         =   "&Screens"
         Begin VB.Menu mnuScreenWorkshop 
            Caption         =   "&Screen Workshop"
         End
      End
      Begin VB.Menu mnuSystemBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "&Reload"
         Begin VB.Menu mnuReloadCommands 
            Caption         =   "&Commands"
         End
         Begin VB.Menu mnuReloadHelp 
            Caption         =   "&Help"
         End
         Begin VB.Menu mnuReloadScreens 
            Caption         =   "&Screens"
         End
         Begin VB.Menu mnuReloadBar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReloadSystem 
            Caption         =   "&System"
         End
      End
   End
   Begin VB.Menu mnuPlayersMain 
      Caption         =   "&Players"
      Begin VB.Menu mnuViewPlayers 
         Caption         =   "&View"
      End
      Begin VB.Menu mnuPlayerManage 
         Caption         =   "&Manage"
      End
   End
   Begin VB.Menu mnuAreasMain 
      Caption         =   "&Areas"
      Begin VB.Menu mnuAreaCreate 
         Caption         =   "&Create"
      End
      Begin VB.Menu mnuAreaView 
         Caption         =   "&View"
      End
      Begin VB.Menu mnuAreaManage 
         Caption         =   "&Manage"
      End
   End
   Begin VB.Menu mnuRoomsMain 
      Caption         =   "&Rooms"
      Begin VB.Menu mnuRoomCreate 
         Caption         =   "&Create"
      End
      Begin VB.Menu mnuRoomView 
         Caption         =   "&View"
      End
      Begin VB.Menu mnuRoomManage 
         Caption         =   "&Manage"
      End
   End
   Begin VB.Menu mnuObjectsMain 
      Caption         =   "&Objects"
   End
   Begin VB.Menu mnuNPCMain 
      Caption         =   "&NPCs"
   End
   Begin VB.Menu mnuQuestsMain 
      Caption         =   "&Quests"
   End
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' NOTES:
'===========================================================================================
'
Public iConnects As Integer
Public iSockets As Integer
Public sRequestID As String
Dim WithEvents MudSockets As Winsock
Attribute MudSockets.VB_VarHelpID = -1
Public TheHelp As New clsHelp

'===========================================================================================
' Sub:      MDIForm_Load()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Setup the connections to the database(s).
'           Load system settings and data.
'           Initialize data-structures in the HUB
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Private Sub MDIForm_Load()

    Set MyConn = New ADODB.Connection
    Set MyRS = New ADODB.Recordset
    Set MyRS2 = New ADODB.Recordset
    
    Set MudSys = New clsSystem
    
    Set HUB.Areas = New clsAreas
    Set HUB.ThePlayers = New clsPlayers

    Me.Show
    Me.Caption = Me.Caption & " - Listening"
    Winsock1(0).LocalPort = 4040
    Winsock1(0).Listen
    Load frmLog
    frmLog.Show
    
    AddToLog "Connecting to System Database..."
    MyConn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\data\MafiaMUD.mdb"
    MyConn.open
    DoEvents
    AddToLog "Connected!"
    AddToLog " "
    
    AddToLog "Loading ANSI Codes..."
    ANSICodes.Set_ANSI_CODES
    AddToLog "  - Load Successful"
    
    AddToLog "Loading Settings..."
    LoadSystem
    AddToLog "  - Load Successful"
    
    AddToLog "Loading Screens..."
    LoadScreens
    AddToLog "  - Load Successful"
    
    AddToLog "Loading Help System..."
    LoadHelp
    AddToLog "  - Load Successful"
    
    AddToLog "Loading Commands..."
    LoadCommands
    AddToLog "  - Load Successful"
    
    AddToLog "Loading Areas..."
    LoadAreas
    AddToLog "  - Loading Rooms..."
    
    '=======================================================================================
    ' ToDo: Get related code changed over to the new area->room/player->area->room structure
    '       (See below)
    '=======================================================================================
    LoadRooms
    
    '=======================================================================================
    ' ToDo: Get related code changed over to the new area->room/player->area->room structure
    '=======================================================================================
    NewLoadRooms
    
    AddToLog "    - Load Successful"
    DoEvents
    LoadExits
    NewLoadExits
    AddToLog "   - Loading Objects..."
    AddToLog "   - Loading Mobiles..."
    AddToLog "   - Loading NPCs..."
    AddToLog " "
    
    '=======================================================================================
    ' Note: Consideration is being made to seperate the player database from the rest of the
    '       system and game data. Significant work would need to be done to make it happen.
    '=======================================================================================
    AddToLog "Connecting to Player Database"
    AddToLog "Connected!"
    
    AddToLog " "
    AddToLog "System Ready on "
    AddToLog "   - Address : " & Winsock1(0).LocalIP & " - Port : " & Winsock1(0).LocalPort
    
End Sub


'===========================================================================================
' Sub:      mnuScreensReload_Click()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-31-2012
'===========================================================================================
' Descript: Handle the menu item click event to reload the screen data while the game is
'           live.
'
'===========================================================================================
' Notes:    This routine has not been tested under load or with a respectively large amount
'           amount of data. It is possible that a user could experience a screen not being
'           shown during that time. <BrianKnust@gmail.com> <05-21-2012>
'
'           5-31-2012: Deprecated this sub-routine due to menu changes. See below for the
'           new menu item click handler.
'===========================================================================================
'
'Private Sub mnuScreensReload_Click()
'    AddToLog "Reloading ANSI Screens"
'    frmParent.LoadScreens
'    AddToLog "  - Reload Successful"
'End Sub


'===========================================================================================
' Sub:      mnuReloadScreens_Click()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-32-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-31-2012
'===========================================================================================
' Descript: Handle the menu item click event to reload the screen data.
'
'===========================================================================================
' Notes:
'===========================================================================================
'
Private Sub mnuReloadScreens_Click()
    AddToLog "Reloading ANSI Screens"
    frmParent.LoadScreens
    AddToLog "  - Reload Successful"
End Sub


'===========================================================================================
' Sub:      mnuScreenWorkship_Click()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handle the menu item click event to open the screen workshop.
'
'===========================================================================================
' Notes:    Currently the screen workshop project has been put on hold. PabloDraw from
'           Pecoe software is highly recommended for use to design game ANSI graphics
'
'===========================================================================================
'
Private Sub mnuScreenWorkshop_Click()
    Load frmScreenShop
    frmScreenShop.Show
End Sub


Private Sub mnuSysCommands_Click()
    frmCommands.Show
End Sub

'===========================================================================================
' Sub:      Winsock1_Close (Integer)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handle the Winsock Close event when a user disconnects from the server
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Private Sub Winsock1_Close(Index As Integer)
'ToDo:
    Winsock1(Index).Close
    iConnects = iConnects - 1
    sbMud.Panels(1).Text = iConnects & ":Cons/" & Winsock1.UBound & ":Socks"
End Sub


'===========================================================================================
' Sub:      Winsock1_Connect (Integer)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handle the Winsock Connect event. (Currently not being used)
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Private Sub Winsock1_Connect(Index As Integer)
'ToDo:
End Sub


'===========================================================================================
' Sub:      Winsock1_ConnectionRequest (Integer, ByVal Long)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handle the Winsock ConnectionRequest Event when a player attempts to connect to
'           the server. Check through the existing array of connections for an open socket,
'           and if none exist create a new one at the end of the array.
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'ToDo:
Dim i As Integer
Dim j As Integer
    
    
    If Index = 0 Then
        sRequestID = requestID
        'iConnects = iConnects + 1
        For i = 1 To Winsock1.UBound
            If Not Winsock1(i).State = sckConnected Then
                'Load Winsock1(i)
                Winsock1(i).LocalPort = 4040
                Winsock1(i).Accept sRequestID
                iConnects = iConnects + 1
                HUB.Players(i).PlrIndex = i
                HUB.Players(i).PlrState = PLR_CONNECTED
        
                sbMud.Panels(1).Text = iConnects & ":Cons/" & Winsock1.UBound & "Sock"
                AddToLog "Connection request by IP:" & Winsock1(i).RemoteHostIP & " on " & Date & " " & Time
                
                ShowScreenToUser i, "Intro1", True
                
                Exit Sub
            End If
        Next i
        iConnects = iConnects + 1
        Load Winsock1(iConnects)
        Winsock1(iConnects).LocalPort = 4040
        Winsock1(iConnects).Accept sRequestID
        ReDim Preserve HUB.Players(iConnects)
        Set HUB.Players(iConnects) = New clsPlayer
        HUB.Players(iConnects).PlrIndex = iConnects
        HUB.Players(iConnects).PlrState = PLR_CONNECTED
        AddToLog "Connection request by IP:" & Winsock1(iConnects).RemoteHostIP & " on " & Date & " " & Time
        
        ShowScreenToUser iConnects, "Intro1", True
        
        sbMud.Panels(1).Text = iConnects & ":Cons/" & Winsock1.UBound & "Sock"
    End If
End Sub


'===========================================================================================
' Sub:      Winsock1_DataArrival (Integer, ByVal Long)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handle the Winsock DataArrival Event triggered when data arrives on a socket
'           in the array.
'
'           Get the data from the sockect and process it. This data is usually user names,
'           passwords, and commands.
'
'===========================================================================================
' Notes:    Data coming is not secure, players and administrators are encouraged to use
'           unique passwords that are not used for secure applications such as online
'           banking and other such accounts.
'
'===========================================================================================
'
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'ToDo:
    Dim ItemData As String
    Dim plrName As String
    Dim crlf As Integer
    Winsock1(Index).GetData ItemData, vbString
    
    
    If ItemData = vbNullString Then Exit Sub
    
    If InStr(1, ItemData, vbCrLf) > 1 Then
        ItemData = mID$(ItemData, 1, Len(ItemData) - 2)
    End If
    
    ParseText Index, ItemData
    
End Sub


'===========================================================================================
' Sub:      Winsock1_Error (Integer, ByVal Integer, String, ByVal Long, ByVal String
'                           ByVal String, ByVal String, ByVal Long, Boolean)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handles the Winsock Error event for the sockets array. (Currently not used)
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long _
, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'ToDo:
End Sub


'===========================================================================================
' Sub:      Winsock1_SendComplete (Integer)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handles the Winsock SendComplete Event (Currently Not Used)
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Private Sub Winsock1_SendComplete(Index As Integer)
'ToDo:
End Sub


'===========================================================================================
' Sub:      Winsock1_SendProgress (Integer, ByVal Long, ByVal Long)
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Handles the Winsock SendProgess Method (Currently Not Used)
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Private Sub Winsock1_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'ToDo:
End Sub


'===========================================================================================
' Sub:      LoadScreens ()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Load the screen information from the database.
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub LoadScreens()
Dim i As Integer
    With MyRS
    .open "SELECT * FROM SCREENS", MyConn, adOpenStatic, adLockOptimistic, adCmdText
    .MoveFirst
    If .RecordCount >= 1 Then
        For i = 1 To .RecordCount
            ReDim Preserve HUB.Screens(i)
            Set HUB.Screens(i) = New clsScreens
            If !SCREEN_DATA <> vbNull Then
                HUB.Screens(i).ScreenData = !SCREEN_DATA
            End If
            HUB.Screens(i).ScreenID = !SCREEN_ID
            HUB.Screens(i).ScreenLevel = !SCREEN_LEVEL
            HUB.Screens(i).ScreenName = !SCREEN_NAME
            HUB.Screens(i).IS_File = !FROMFILE
            If !FROMFILE = True Then
                HUB.Screens(i).SFileName = !SFileName
            End If
            .MoveNext
        Next i
    End If
    .Close
    End With
End Sub


'===========================================================================================
' Sub:      LoadHelp ()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Load the Help information from the database.
'
'===========================================================================================
' Notes:    The Help system is idealy used from files, rather than storing potentially
'           dozens or hundreds of screens in memory. A tradeoff is made on performance for
'           preservation of system memory for use in other tasks.
'
'===========================================================================================
'

Public Sub LoadHelp()
Dim i As Integer
Dim HN As String
Dim FF As Boolean
Dim HFN As String
Dim HS As String
Dim HSC As String
    
    With MyRS
    .open "SELECT * FROM HELPSYS", MyConn, adOpenStatic, adLockOptimistic, adCmdText
    .MoveFirst
    If .RecordCount >= 1 Then
        For i = 1 To .RecordCount
            HN = !HELPNAME
            FF = !FROMFILE
            HFN = !HFileName
            If Not (!HELP_SYNTAX = vbNull) Then
                HS = !HELP_SYNTAX
            Else
                HS = ""
            End If
            If Not (!HELP_SCREEN = vbNull) Then
                HSC = !HELP_SCREEN
            Else
                HSC = ""
            End If
            
            If Not (!HFileName = vbNull) Then
                HFN = !HFileName
            Else
                HFN = ""
            End If
            
            'TheHelp.Helps.Add !HELPNAME, !FROMFILE, !HFileName, !HELP_SYNTAX, !HELP_SCREEN
            TheHelp.Helps.Add HN, FF, HFN, HS, HSC, HN
            .MoveNext
        Next i
    End If
    .Close
    End With
End Sub


'===========================================================================================
' Sub:      LoadSystem ()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Load the system information/defaults
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub LoadSystem()
Dim i As Integer
    With MyRS
       .open "SELECT * FROM SYSTEM", MyConn, adOpenStatic, adLockOptimistic, adCmdText
        .MoveFirst
        If .RecordCount >= 1 Then
            HUB.MudSys.Build = !SYSBUILD
            HUB.MudSys.DamMVP = !SYSDAMMVP
            HUB.MudSys.DamPVM = !SYSDAMPVM
            HUB.MudSys.DamPVP = !SYSDAMPVP
            HUB.MudSys.DodgeMod = !SYSDODGEMOD
            HUB.MudSys.MaxLevel = !SYSMAXLEVEL
            HUB.MudSys.MudName = !SYSMUDNAME
            HUB.MudSys.ParryMod = !SYSPARRYMOD
            HUB.MudSys.PKLoot = !SYSPKLOOT
            HUB.MudSys.Proto = !SYSPROTO
            HUB.MudSys.StunPVP = !SYSSTUNPVP
            HUB.MudSys.StunReg = !SYSSTUNREG
            HUB.MudSys.WaitForAuth = !SYSWAITFORAUTH
        
        End If
        .Close
    End With
End Sub


'===========================================================================================
' Sub:      LoadCommands ()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Load the user/player commands
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub LoadCommands()
Dim i As Integer

    With MyRS
        .open "SELECT * FROM COMMANDS", MyConn, adOpenStatic, adLockOptimistic, adCmdText
        .MoveFirst
        
        If .RecordCount >= 1 Then
            For i = 1 To .RecordCount
                ReDim Preserve HUB.Commands(i)
                Set HUB.Commands(i) = New clsCommands
                HUB.Commands(i).Name = !CMDNAME
                HUB.Commands(i).Code = !CMDCODE
                HUB.Commands(i).Pos = !CMDPOS
                HUB.Commands(i).Level = !CMDLEVEL
                HUB.Commands(i).Log = !CMDLOG
            .MoveNext
            Next i
        End If
    .Close
    End With

End Sub


'===========================================================================================
' Sub:      LoadAreas ()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-24-2012
'===========================================================================================
' Descript: Load the areas and rooms
'
'===========================================================================================
' Notes: 5-24-2011 <Brian Knust> <BrianKnust@gmail.com>
'        See ToDo: Note below in the code.
'
'===========================================================================================
'
Public Sub LoadAreas()
Dim i, k As Integer
Dim TheArea As clsArea
Dim theRoom As clsRoom

    Set TheArea = New clsArea
    Set theRoom = New clsRoom
    MyRS.open "SELECT * FROM AREAS", MyConn, adOpenStatic, adLockOptimistic, adCmdText
    MyRS.MoveFirst
    
    If MyRS.RecordCount >= 1 Then
        For i = 1 To MyRS.RecordCount
            TheArea.AName = MyRS!AName
            TheArea.HVnum = MyRS!HVnum
            TheArea.LVnum = MyRS!LVnum
            TheArea.ResMessage = MyRS!RESET_MESSAGE
            TheArea.ResTime = MyRS!RESET_TIMER
            
            Areas.Add TheArea, MyRS!AIndex
            
            MyRS2.open "SELECT * From Rooms WHERE (((ROOMS.AINDEX)=" & MyRS!AIndex & "));", MyConn, adOpenStatic, adLockOptimistic, adCmdText
            MyRS2.MoveFirst
            
            If MyRS2.RecordCount >= 1 Then
                For k = 1 To MyRS2.RecordCount
                    
                    
                    '=======================================================================
                    '
                    '=======================================================================
                    ' ToDo: - Add Code to add rooms to areas. (Complete, see NewLoadRooms)
                    '       - Test Code thouroughly (step through and make sure all data is
                    '         in place properly)
                    '       - Begin rewriting code for room access and descriptions
                    '       - Exits
                    '=======================================================================
                    
                    
                    MyRS2.MoveNext
                Next k
            End If
            MyRS2.Close
            MyRS.MoveNext
        Next i
    End If
    MyRS.Close
End Sub


'===========================================================================================
' Sub:      LoadRooms ()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Load the rooms (Currently not loaded by areas)
'
'===========================================================================================
' Notes:
'
'===========================================================================================
'
Public Sub LoadRooms()
Dim i As Integer
    With MyRS
        .open "SELECT * FROM ROOMS", MyConn, adOpenStatic, adLockOptimistic, adCmdText
        .MoveFirst
        
        If .RecordCount >= 1 Then
        
            For i = 1 To .RecordCount
                ReDim Preserve HUB.Rooms(i)
                Set HUB.Rooms(i) = New clsRooms
                    HUB.Rooms(i).LongDesc = !ROOMLONG
                    HUB.Rooms(i).Name = !ROOMNAME
                    HUB.Rooms(i).RoomIndex = !ROOMINDX
                    HUB.Rooms(i).ShortDesc = !ROOMSHORT
                    HUB.Rooms(i).VNum = !ROOMVNUM
                .MoveNext
            Next i
        End If
        .Close
    End With
End Sub

Public Sub NewLoadRooms()
Dim i As Integer
Dim theRoom As clsRoom
    
    Set theRoom = New clsRoom
    With MyRS
        .open "SELECT * FROM ROOMS", MyConn, adOpenStatic, adLockOptimistic, adCmdText
        .MoveFirst
        
        If .RecordCount >= 1 Then
            For i = 1 To .RecordCount
                theRoom.AIndex = !AIndex
                theRoom.RMIndex = !ROOMINDX
                theRoom.RName = !ROOMNAME
                theRoom.VNum = !ROOMVNUM
                theRoom.LongDesc = !ROOMLONG
                theRoom.ShortDesc = !ROOMSHORT
                
                Areas.Item(!AIndex).Rooms.Add theRoom, !ROOMVNUM
                                
                .MoveNext
            Next i
        End If
        .Close
    End With
End Sub

'===========================================================================================
' Sub:      LoadExits()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     05-20-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 05-21-2012
'===========================================================================================
' Descript: Load the exit (room-to-room) data.
'
'===========================================================================================
' Notes:    Currently exits are loaded seperately from rooms and areas.
'
'===========================================================================================
'
Public Sub LoadExits()
Dim i As Integer
Dim r As Integer

    'Use the RecordSet Object 'MyRs'
    With MyRS
        
        'Open The RecordSet
        .open "SELECT * FROM EXITS", MyConn, adOpenStatic, adLockOptimistic, adCmdText
        
        'Move to the first object in the Recordset
        .MoveFirst
        
        'Is there at least 1 (one) record in the RecordSet?
        If .RecordCount >= 1 Then
        
            'Let's loop through the records and..
            For i = 1 To .RecordCount
                
                '..find the room that the exit belongs to..
                For r = 1 To UBound(Rooms)
                    If Rooms(r).VNum = !EXROOMFROM Then Exit For
                Next r
                
                '..and then add the exit to the room
                Rooms(r).AddExit !EXROOMTO, !EXDIR, False, False, False
                '.AddExit        RoomGoing, Direct, Door?, Clmb?, Fly?
                
                'Go to the next Record
                .MoveNext
                
            'Do it again
            Next i
        End If
        .Close
    End With
End Sub


'===========================================================================================
' Sub:      NewLoadExits()
' Coded:    Brian Knust <BrianKnust@gmail.com>
' Date:     06-06-2012
' Mod-By:   Brian Knust <BrianKnust@gmail.com>
' Mod-Date: 06-06-2012
'===========================================================================================
' Descript: Load the exit (room-to-room) data.
'
'===========================================================================================
' Notes:    Currently exits are loaded seperately from rooms and areas.
'
'           It was suggested that the exit in the opposite direction be added at the same
'           time, but this is not recommended because not all exits may be bi-directional.
'           During building, a return exit will be automatically created, unless specified
'           in the command, or removed by the builder.
'
'===========================================================================================
'
Public Sub NewLoadExits()
Dim i, r As Integer
Dim theExit As clsExit
Set theExit = New clsExit
    
    With MyRS
    
        .open "SELECT * FROM EXITS", MyConn, adOpenStatic, adLockOptimistic, adCmdText
        .MoveFirst
        
        If .RecordCount >= 1 Then
            For i = 1 To .RecordCount
                theExit.AreaFrom = !EXAREAFROM
                theExit.AreaTo = !EXAREATO
                theExit.Climb = !EXCLIMB
                theExit.Direction = !EXDIR
                theExit.Door = !EXDOOR
                theExit.Fly = !EXFLY
                If Not (!EXKEYWORD = vbNull) Then
                    theExit.KeyWord = !EXKEYWORD
                Else
                    theExit.KeyWord = ""
                End If
                theExit.RoomFrom = !EXROOMFROM
                theExit.RoomTo = !EXROOMTO
                theExit.Swim = !EXSWIM
                
                Areas.Item(theExit.AreaFrom).Rooms.Item(theExit.RoomFrom).Exits.Add theExit, !ExIndex
                
                .MoveNext
            Next i

        End If
        .Close
    End With
End Sub
