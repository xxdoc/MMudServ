Attribute VB_Name = "Tables"
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
Public Sub SkillTable(ByVal Index As Integer, ByVal Skill As String, ByVal DataText As String)
On Error Resume Next
    Dim i As Integer
    
    Skill = UCase$(Skill)
    
    For i = 1 To (UBound(HUB.Commands) - 1)
        If UCase$(HUB.Commands(i).Name) = Skill Then
            Exit For
        End If
    Next i

    With HUB.Commands(i)
        If UCase$(.Name) <> Skill Then
            SendToUser Index, "%BUnknown Command%n" & vbCrLf
            Exit Sub
        End If
        
        If .Code = "do_cocreate" Then
            DoFuns.DoCoCreate Index, DataText
        ElseIf .Code = "do_coedit" Then
            DoFuns.DoCoEdit Index, DataText
        ElseIf .Code = "do_coelock" Then
            DoFuns.DoCOELock Index, DataText
        ElseIf .Code = "do_create" Then
            DoCreate Index, DataText
        ElseIf .Code = "do_east" Then           'East
            DoEast Index, DataText
        ElseIf .Code = "do_hcreate" Then
            DoFuns.DoHCreate Index, DataText
        ElseIf .Code = "do_hedit" Then
            DoFuns.DoHEdit Index, DataText
        ElseIf .Code = "do_helock" Then
            DoFuns.DoHELock Index, DataText
        ElseIf .Code = "do_help" Then           'Help
            DoHelp Index, DataText
        ElseIf .Code = "do_hire" Then
            DoFuns.DoHire Index, DataText
        ElseIf .Code = "do_mcreate" Then
            DoFuns.DoMCreate Index, DataText
        ElseIf .Code = "do_medit" Then
            DoFuns.DoMEdit Index, DataText
        ElseIf .Code = "do_melock" Then
            DoFuns.DoMELock Index, DataText
        ElseIf .Code = "do_north" Then          'North
            DoNorth Index, DataText
        ElseIf .Code = "do_northeast" Then      'NorthEast
            DoNorthEast Index, DataText
        ElseIf .Code = "do_northwest" Then      'NorthWest
            DoNorthWest Index, DataText
        ElseIf .Code = "do_ocreate" Then
            DoFuns.DoOCreate Index, DataText
        ElseIf .Code = "do_oedit" Then
            DoFuns.DoOEdit Index, DataText
        ElseIf .Code = "do_oelock" Then
            DoFuns.DoOELock Index, DataText
        ElseIf .Code = "do_rcreate" Then
            DoFuns.DoRCreate Index, DataText
        ElseIf .Code = "do_redit" Then
            DoFuns.DoREdit Index, DataText
        ElseIf .Code = "do_relock" Then
            DoFuns.DoRELock Index, DataText
        ElseIf .Code = "do_reload" Then         'Reload
            DoReload Index, DataText
        ElseIf .Code = "do_say" Then            'Say
            DoSay Index, DataText
        ElseIf .Code = "do_sell" Then           'Sell
            DoSell Index, DataText
        ElseIf .Code = "do_score" Then          'Score
            DoScore Index, DataText
        ElseIf .Code = "do_shoot" Then          'Shoot
            DoShoot Index, DataText
        ElseIf .Code = "do_shout" Then          'Shout
            DoShout Index, DataText
        ElseIf .Code = "do_south" Then          'South
            DoSouth Index, DataText
        ElseIf .Code = "do_southeast" Then      'Southeast
            DoSouthEast Index, DataText
        ElseIf .Code = "do_southwest" Then      'Southwest
            DoSouthWest Index, DataText
        ElseIf .Code = "do_yell" Then           'Yell
            DoYell Index, DataText
        ElseIf .Code = "do_hire" Then           'Hire
            DoHire Index, DataText
        ElseIf .Code = "do_screen" Then         'Screen
            DoScreen Index, DataText
        ElseIf .Code = "do_west" Then           'West
            DoWest Index, DataText
        End If
    End With
    
End Sub

'|---------|---------|---------|---------|---------|---------|---------|---------
