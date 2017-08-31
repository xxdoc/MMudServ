Attribute VB_Name = "modScreens"
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
Public Sub ShowScreenToUser(ByVal vIndex As Integer, ByVal vScreen As String, Optional ByVal vClear As Boolean)
Dim i As Integer
Dim TheText As String
    For i = 1 To UBound(HUB.Screens)
        If UCase$(vScreen) = UCase$(HUB.Screens(i).ScreenName) Then
            If HUB.Screens(i).IS_File = False Then
                If vClear = True Then
                    SendToUser vIndex, cl
                    SendToUser vIndex, HUB.Screens(i).ScreenData
                Else
                    SendToUser vIndex, HUB.Screens(i).ScreenData
                End If
            Else
                If UCase$(vScreen) = UCase$(HUB.Screens(i).ScreenName) Then
                    If vClear = True Then
                        SendToUser vIndex, cl
                        Open App.Path & "\data\screens\" & HUB.Screens(i).SFileName For Input As #1 'ToDo: Write Send From File Code
                        Do While Not EOF(1)
                            Input #1, TheText
                            SendToUser vIndex, TheText
                        Loop
                        Close #1
                    Else
                        Open App.Path & "\data\screens\" & HUB.Screens(i).SFileName For Input As #1 'ToDo: Write Send From File Code
                        Do While Not EOF(1)
                            Input #1, TheText
                            SendToUser vIndex, TheText
                        Loop
                        Close #1
                        'ToDo: Write Send From File Code
                    End If
                End If
            End If
            Exit For
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
Public Sub ShowScreenFromFile(ByVal vIndex As Integer, ByVal vScreen As String, Optional ByVal vClear As Boolean)
Dim i As Integer
Dim FreeFile As Integer
Dim TheText As String


    For i = 1 To UBound(HUB.Screens)
        If UCase$(vScreen) = UCase$(HUB.Screens(i).ScreenName) Then
            If vClear = True Then
                SendToUser vIndex, cl
                Open App.Path & "\data\screens\" & HUB.Screens(i).SFileName For Input As #1 'ToDo: Write Send From File Code
                    Do While Not EOF(1)
                        Input #1, TheText
                        SendToUser vIndex, TheText
                    Loop
                Close #1
            Else
                Open App.Path & "\data\screens\" & HUB.Screens(i).SFileName For Input As #1 'ToDo: Write Send From File Code
                    Do While Not EOF(1)
                        Input #1, TheText
                        SendToUser vIndex, TheText
                    Loop
                Close #1
                'ToDo: Write Send From File Code
            End If
        End If
    Next i

End Sub
