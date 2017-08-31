Attribute VB_Name = "ANSICodes"
Option Explicit

Public k0 As String
Public r0 As String
Public g0 As String
Public y0 As String
Public b0 As String
Public p0 As String
Public c0 As String
Public w0 As String
Public fRes0 As String
    
Public k1 As String
Public r1 As String
Public g1 As String
Public y1 As String
Public b1 As String
Public p1 As String
Public c1 As String
Public w1 As String
Public fRes1 As String

Public bk0 As String
Public br0 As String
Public bg0 As String
Public by0 As String
Public bb0 As String
Public bp0 As String
Public bc0 As String
Public bw0 As String
Public bRes0 As String
    
Public bk1 As String
Public br1 As String
Public bg1 As String
Public by1 As String
Public bb1 As String
Public bp1 As String
Public bc1 As String
Public bw1 As String
Public bRes1 As String

Public cl As String


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
Public Sub Set_ANSI_CODES()
Dim CSI As String
CSI = Chr$(27) + Chr$(91)
    cl = CSI + "40m" & CSI + "2J" + CSI + "40m"

    k0 = CSI + "0;30m"
    r0 = CSI + "0;31m"
    g0 = CSI + "0;32m"
    y0 = CSI + "0;33m"
    b0 = CSI + "0;34m"
    p0 = CSI + "0;35m"
    c0 = CSI + "0;36m"
    w0 = CSI + "0;37m"
    fRes0 = CSI + "0;39m"

    k1 = CSI + "1;30m"
    r1 = CSI + "1;31m"
    g1 = CSI + "1;32m"
    y1 = CSI + "1;33m"
    b1 = CSI + "1;34m"
    p1 = CSI + "1;35m"
    c1 = CSI + "1;36m"
    w1 = CSI + "1;37m"
    fRes1 = CSI + "1;39m"
    
    bk0 = CSI + "0;40m"
    br0 = CSI + "0;41m"
    bg0 = CSI + "0;42m"
    by0 = CSI + "0;43m"
    bb0 = CSI + "0;44m"
    bp0 = CSI + "0;45m"
    bc0 = CSI + "0;46m"
    bw0 = CSI + "0;47m"
    bRes0 = CSI + "0;49m"

    bk1 = CSI + "1;40m"
    br1 = CSI + "1;41m"
    bg1 = CSI + "1;42m"
    by1 = CSI + "1;43m"
    bb1 = CSI + "1;44m"
    bp1 = CSI + "1;45m"
    bc1 = CSI + "1;46m"
    bw1 = CSI + "1;47m"
    bRes1 = CSI + "1;49m"
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
Public Function ParseANSI(ByVal DataText) As String
Dim FG As String
Dim BG As String
Dim CSI As String
Dim ANSIs() As String
Dim i As Integer
ParseANSI = DataText

CSI = Chr$(27) & Chr$(91)
FG = "%"
BG = "^"

ReDim Preserve ANSIs(1)
For i = 1 To Len(DataText)
    If InStr(1, "%", mID$(DataText, i, 1)) Or InStr(1, "^", mID$(DataText, i, 1)) Then
        ReDim Preserve ANSIs(UBound(ANSIs) + 1)
        ANSIs(UBound(ANSIs)) = mID$(DataText, i, 2)
        'MsgBox "Found " & ANSIs(UBound(ANSIs))
    End If
Next i

i = 0
For i = 0 To UBound(ANSIs)
    Select Case ANSIs(i)
        Case "%w"
            ParseANSI = Replace(ParseANSI, ANSIs(i), w0, 1, 1)
        Case "%W"
            ParseANSI = Replace(ParseANSI, ANSIs(i), w1, 1, 1)
        Case "%c"
            ParseANSI = Replace(ParseANSI, ANSIs(i), c0, 1, 1)
        Case "%C"
            ParseANSI = Replace(ParseANSI, ANSIs(i), c1, 1, 1)
        Case "%p"
            ParseANSI = Replace(ParseANSI, ANSIs(i), p0, 1, 1)
        Case "%P"
            ParseANSI = Replace(ParseANSI, ANSIs(i), p1, 1, 1)
        Case "%b"
            ParseANSI = Replace(ParseANSI, ANSIs(i), b0, 1, 1)
        Case "%B"
            ParseANSI = Replace(ParseANSI, ANSIs(i), b1, 1, 1)
        Case "%y"
            ParseANSI = Replace(ParseANSI, ANSIs(i), y0, 1, 1)
        Case "%Y"
            ParseANSI = Replace(ParseANSI, ANSIs(i), y1, 1, 1)
        Case "%g"
            ParseANSI = Replace(ParseANSI, ANSIs(i), g0, 1, 1)
        Case "%G"
            ParseANSI = Replace(ParseANSI, ANSIs(i), g1, 1, 1)
        Case "%r"
            ParseANSI = Replace(ParseANSI, ANSIs(i), r0, 1, 1)
        Case "%R"
            ParseANSI = Replace(ParseANSI, ANSIs(i), r1, 1, 1)
        Case "%k"
            ParseANSI = Replace(ParseANSI, ANSIs(i), k0, 1, 1)
        Case "%K"
            ParseANSI = Replace(ParseANSI, ANSIs(i), k1, 1, 1)
        Case "%n"
            ParseANSI = Replace(ParseANSI, ANSIs(i), fRes0, 1, 1)
        Case "%N"
            ParseANSI = Replace(ParseANSI, ANSIs(i), fRes1, 1, 1)
        Case "%%"
            ParseANSI = Replace(ParseANSI, ANSIs(i), "%", 1, 1)
        Case "%e"
            ParseANSI = Replace(ParseANSI, ANSIs(i), cl, 1, 1)
    
    
        Case "^w"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bw0, 1, 1)
        Case "^W"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bw1, 1, 1)
        Case "^c"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bc0, 1, 1)
        Case "^C"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bc1, 1, 1)
        Case "^p"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bp0, 1, 1)
        Case "^P"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bp1, 1, 1)
        Case "^b"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bb0, 1, 1)
        Case "^B"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bb1, 1, 1)
        Case "^y"
            ParseANSI = Replace(ParseANSI, ANSIs(i), by0, 1, 1)
        Case "^Y"
            ParseANSI = Replace(ParseANSI, ANSIs(i), by1, 1, 1)
        Case "^g"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bg0, 1, 1)
        Case "^G"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bg1, 1, 1)
        Case "^r"
            ParseANSI = Replace(ParseANSI, ANSIs(i), br0, 1, 1)
        Case "^R"
            ParseANSI = Replace(ParseANSI, ANSIs(i), br1, 1, 1)
        Case "^k"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bk0, 1, 1)
        Case "^K"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bk1, 1, 1)
        Case "^n"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bRes0, 1, 1)
        Case "^N"
            ParseANSI = Replace(ParseANSI, ANSIs(i), bRes1, 1, 1)
        Case "^^"
            ParseANSI = Replace(ParseANSI, ANSIs(i), "^", 1, 1)
    End Select
Next i

End Function
