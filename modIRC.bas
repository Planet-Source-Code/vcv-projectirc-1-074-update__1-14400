Attribute VB_Name = "modIRC"
'* projectIRC version 1.0
'* By Matt C, sappy@adelphia.net
'* Feel free to EMail me with any questions you may have.

'* Module that handles many of the Client's procedures,
'* Feel free to use in your project as long as you give me proper credit

'/me is using projectIRC $version $+ , on $server port $port me = $me on channel $chan
Global strVersionReply As String
Global path As String
'* Channels and Queries
Global Const MAX_CHANNELS = 30
Global Const MAX_QUERIES = 30
Global Const MAX_DCCCHATS = 30
Global Const MAX_DCCSENDS = 30
Public Channels(1 To MAX_CHANNELS)  As Channel
Public Queries(1 To MAX_QUERIES)    As Query
Public DCCChats(1 To MAX_DCCCHATS)  As DCCChat
Public DCCSends(1 To MAX_DCCSENDS) As DCCSend
Public intChannels  As Integer
Public intQueries   As Integer
Public intDCCChats  As Integer
Public intDCCSends  As Integer

'* Variables for incoming commands
Type ParsedData
    bHasPrefix   As Boolean
    strParams()  As String
    intParams    As Integer
    strFullHost  As String
    strCommand   As String
    strNick      As String
    strIdent     As String
    strHost      As String
    AllParams    As String
End Type

'* ANSI Formatting character values
Global Const bold = 2
Global Const underline = 31
Global Const Color = 3
Global Const REVERSE = 22
Global Const ACTION = 1

'* ANSI Formatting characters
Global strBold As String
Global strUnderline As String
Global strColor As String
Global strReverse As String
Global strAction As String

'Nick storage for nick list inchannels
Type Nick
    Nick    As String '* 40
    Op      As Boolean
    Voice   As Boolean
    Helper  As Boolean
    Host    As String '* 100
    IDENT   As String '* 30
End Type

'Mode storage for each channel
Type typMode
    mode    As String
    bPos    As Boolean
End Type

'* AWAY options
Public lngGoneFor   As Long
Public bAnnounce    As Boolean

Public lngPingReply As Long

Type DCC_INFO
    Id      As String
    File    As String
    Port    As Long
    IP      As String ' * 16
    Nick    As String ' * 40
    type    As Integer
    Size    As Long
End Type

Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209



'* There are all the known IRC Numeric Codes
Global Const IRC_WELCOMEMSG = 1
Global Const IRC_LOCALHOSTIRCD = 2
Global Const IRC_SERVERCREATED = 3
Global Const IRC_AVAILABLE = 5

Global Const STATS_USERS = 251
Global Const STATS_IRCCOP = 252
Global Const STATS_CHANNELS = 254
Global Const STATS_CLIENTSSERVERS = 255
Global Const STATS_SERVERTOOHEAVY = 263
Global Const STATS_LOCALMAXUSERS = 265
Global Const STATS_GLOBALMAXUSERS = 266
Global Const WM_SETFOCUS = &H7

Global Const MODES_UNKNOWNCHARMODE = 501


Public Const LF_FACESIZE = 32
Public Const WM_USER = &H400
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Type CHARFORMAT2
    cbSize As Long
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As String
    bPitchAndFamily As String
    szFaceName As String * LF_FACESIZE
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lcid As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As String
    bAnimation As String
    bRevAuthor As String
    bReserved1 As String
End Type

Public Const SCF_SELECTION = &H1&
Public Const SCF_ALL = &H4&
Public Const EM_SETSEL = &HB1

Function ANSICode(rtf As RichTextBox) As String
    Dim strFinal As String, i As Integer, strChr As String, lngColor As Long
    Dim bBold As Boolean, bUnderline As Boolean, bReverse As Boolean
    lngColor = 0
    
    For i = 0 To Len(rtf.Text) - 1
        strChr = Mid(rtf.Text, i + 1, 1)
        rtf.SelStart = i + 1
        
        'MsgBox strChr & "~b: " & rtf.SelBold & "~u: " & rtf.SelUnderline
        
        If rtf.SelBold = Not bBold Then
            bBold = Not bBold
            strFinal = strFinal & strBold
        End If
        If rtf.SelUnderline = Not bUnderline Then
            bUnderline = Not bUnderline
            strFinal = strFinal & strUnderline
        End If
        If rtf.SelStrikeThru = Not bReverse Then
            bReverse = Not bReverse
            strFinal = strFinal & strReverse
        End If
        If rtf.SelColor <> lngColor Then
            lngColor = rtf.SelColor
            'MsgBox lngColor & "~" & strChr
            strFinal = strFinal & strColor & RAnsiColor(rtf.SelColor)
        End If
        
        strFinal = strFinal & strChr
    Next i
    ANSICode = strFinal
        
End Function


Sub DoTooltip(ByRef strCommand As String, ByRef strInfo As String)
    
    Select Case LCase(strCommand)
        Case "join"
            strCommand = "JOIN <#channel> [key], [#channel2 [key]]"
            strInfo = "Joins the given channels, seperated by a comma."
        Case "part"
            strCommand = "PART <#channel>, [reason]"
            strInfo = "Leaves the given channel, with the given reason (if given any)"
        Case "msg"
            strCommand = "MSG <nick>, <message>"
            strInfo = "Sends a private message to <nick> with <message> as the text."
        Case "query"
            strCommand = "QUERY <nick>, [message]"
            strInfo = "Opens up a query window with <nick> and sends [message] if given."
        Case "voice"
            strCommand = "VOICE <nick>, [[nick2], [nick3], ...]"
            strInfo = "Gives voice status to the given nicks, seperated by a space."
        Case "op"
            strCommand = "OP <nick>, [[nick2], [nick3], ...]"
            strInfo = "Gives operator status to the given nicks, seperated by a space."
        Case "helper"
            strCommand = "HELPER <nick>, [[nick2], [nick3], ...]"
            strInfo = "Gives helper status to the given nicks, seperated by a space."
        Case "whois"
            strCommand = "WHOIS <nick>"
            strInfo = "Requests information on given nicks."
        Case "quit"
            strCommand = "QUIT [message]"
            strInfo = "Closes the connection with the current server, quitting with the given [message], if any."
        Case "nick"
            strCommand = "NICK <newnick>"
            strInfo = "Changes your nick to <newnick>, unless someone else is using it."
        Case "lusers"
            strCommand = "LUSERS"
            strInfo = "Requests user information from the server, including user-count and such."
        Case "id"
            strCommand = "ID <password>"
            strInfo = "Identified with NickServ on servers that support it."
        Case Else
            If AliasExists(strCommand) Then
                strCommand = "/" & UCase(strCommand) & " ???"
                strInfo = "Custom alias, use unknown."
            Else
                strCommand = ""
                strInfo = ""
            End If
    End Select
    
End Sub

Public Function LongIPToIP(dblIP As Double) As String

  ''used with DCC

  Dim arrIP(1 To 4) As Integer

  arrIP(4) = dblIP# Mod 256
  dblIP# = Int(dblIP# / 256)
  arrIP(3) = dblIP# Mod 256
  dblIP# = Int(dblIP# / 256)
  arrIP(2) = dblIP# Mod 256
  dblIP# = Int(dblIP# / 256)
  arrIP(1) = Int(dblIP#)

  LongIPToIP = arrIP(1) & "." & arrIP(2) & "." & arrIP(3) & "." & arrIP(4)

End Function

Public Sub ActionAll(strText As String)
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then
            PutData Channels(i).DataIn, strColor & "06" & strMyNick & " " & strText
            Client.SendData "PRIVMSG " & Channels(i).strName & " :" & strAction & "ACTION " & strText & strAction
        End If
    Next i
End Sub


Public Function BytesToChars(ByVal lngBytes As Long) As String
     Dim strBuffer As String
     strBuffer$ = chr$(lngBytes& Mod 256)
     lngBytes& = Int(lngBytes& / 256)
     strBuffer$ = chr$(lngBytes& Mod 256) & strBuffer$
     lngBytes& = Int(lngBytes& / 256)
     strBuffer$ = chr$(lngBytes& Mod 256) & strBuffer$
     lngBytes& = Int(lngBytes& / 256)
     strBuffer$ = chr$(lngBytes&) & strBuffer$
     
     BytesToChars = strBuffer$
End Function

Function Duration(lngSeconds As Long) As String
    If lngSeconds = 0 Then Duration = "0 seconds": Exit Function
    Dim yrs As Long, mnths As Long, dys As Long, hrs As Long, mns As Long, scs As Long
    
    On Error Resume Next
    yrs = Int(lngSeconds / (60 * 60 * 24 * 365)) '# of seconds in a year
    
    'yrs = 0
    lngSeconds = lngSeconds Mod (60 * 60 * 24 * 365)
    
    mnths = lngSeconds / (60 * 60 * 24 * 12)    '# of seconds in month
    lngSeconds = lngSeconds Mod (60 * 60 * 24 * 12)
    
    dys = lngSeconds / (60 * 60 * 24)    '# of seconds in a day
    lngSeconds = lngSeconds Mod 86400
    
    hrs = lngSeconds / 3600     '# of seconds in an hour
    lngSeconds = lngSeconds Mod 3600
    
    mns = Int(lngSeconds / 60)
    scs = lngSeconds Mod 60
    
    Duration = LeftR(SV(yrs, "year") & _
               SV(mnths, "month") & _
               SV(dys, "day") & _
               SV(hrs, "hour") & _
               SV(mns, "minute") & _
               SV(scs, "second"), 1)
    
               
End Function

Sub EchoActive(strText As String, Optional intColor As Integer = 1)
    On Error Resume Next
    PutData Client.ActiveForm.DataIn, strColor & intColor & strText
    If Err Then PutData Status.DataIn, strColor & intColor & strText
End Sub

Function GetDCCChatIndex(strNick As String) As Integer
    Dim i As Integer

    For i = 1 To intDCCChats
        If LCase(DCCChats(i).strNick) = LCase(strNick) Then
            GetDCCChatIndex = i
            Exit Function
        End If
    Next i
    GetDCCChatIndex = -1
End Function

Function GetDCCSendIndex(strID As String)
    Dim i As Integer

    For i = 1 To intdccsemds
        If LCase(DCCSends(i).Id) = LCase(strID) Then
            GetDCCSendIndex = i
            Exit Function
        End If
    Next i
    GetDCCSendIndex = -1

End Function

Function LongIp(strIP As String) As String
    Dim strNums() As String
    strNums = Split(strIP, ".")
    
    LongIp = (CInt(strNums(0)) * 16777216) + (CInt(strNums(1)) * 65536) + (CInt(strNums(2)) * 256) + (CInt(strNums(3)) * 1)
End Function

Public Function NewDCCChat(strNick As String, strHost As String, lngPort As Long)
    
    'i = GetDCCSendIndex('')
    'If i <> -1 Then
    '    DCCSends(i).SetFocus
    '    Exit Sub
    'End If
    'MsgBox lngPort

    For i = 1 To intDCCChats
        If DCCChats(i).Tag = "" Then
            DCCChats(i).strNick = strNick
            DCCChats(i).strHost = strHost
            DCCChats(i).lngPort = lngPort
            DCCChats(i).Caption = "DCC Chat - " & strNick
            'DCCChats(i).sock.LocalPort = lngPort
            'DCCChats(i).sock.RemotePort = lngPort
            'DCCChats(i).sock.RemoteHost = strHost
            DCCChats(i).lblHost = strHost
            DCCChats(i).lblNick = strNick
            DCCChats(i).Tag = i
            DCCChats(i).Show
            NewDCCChat = i
        End If
    Next i
    
    intDCCChats = intDCCChats + 1
    Set DCCChats(intDCCChats) = New DCCChat
    DCCChats(intDCCChats).strNick = strNick
    DCCChats(intDCCChats).strHost = strHost
    DCCChats(intDCCChats).lngPort = lngPort
    DCCChats(intDCCChats).Caption = "DCC Chat - " & strNick
    'DCCChats(intDCCChats).sock.LocalPort = lngPort
    'DCCChats(intDCCChats).sock.RemotePort = lngPort
    'DCCChats(intDCCChats).sock.RemoteHost = strHost
    DCCChats(intDCCChats).lblHost = strHost
    DCCChats(intDCCChats).lblNick = strNick
    DCCChats(intDCCChats).Show
    DCCChats(intDCCChats).Tag = intDCCChats
    NewDCCChat = intDCCChats
End Function



Public Function NewDCCSend(typDCC As DCC_INFO)
    
    i = GetDCCSendIndex(typDCC.Id)
    'If i <> -1 Then
    '    DCCSends(i).SetFocus
    '    Exit Sub
    'End If

    For i = 1 To intDCCSends
        If DCCSends(i).Id = "" Then
            DCCSends(i).Id = typDCC.Id
            DCCSends(i).lngRemotePort = typDCC.Port
            DCCSends(i).strFile = typDCC.File
            DCCSends(i).strNick = typDCC.Nick
            DCCSends(i).intDCCType = typDCC.type
            DCCSends(i).lngFileSize = typDCC.Size
            DCCSends(i).sock.RemoteHost = typDCC.IP
            DCCSends(i).sock.RemotePort = typDCC.Port
            DCCSends(i).UpdateInfo
            DCCSends(i).Tag = i
            DCCSends(i).Show
            NewDCCSend = i
            Exit Function
        End If
    Next i
    
    intDCCSends = intDCCSends + 1
    Set DCCSends(intDCCSends) = New DCCSend
    DCCSends(intDCCSends).Id = typDCC.Id
    DCCSends(intDCCSends).lngRemotePort = typDCC.Port
    DCCSends(intDCCSends).strFile = typDCC.File
    DCCSends(intDCCSends).strNick = typDCC.Nick
    DCCSends(intDCCSends).intDCCType = typDCC.type
    DCCSends(intDCCSends).lngFileSize = typDCC.Size
    DCCSends(intDCCSends).sock.RemoteHost = typDCC.IP
    DCCSends(intDCCSends).sock.RemotePort = typDCC.Port
    DCCSends(intDCCSends).UpdateInfo
    DCCSends(intDCCSends).Show
    DCCSends(intDCCSends).Tag = intDCCSends
    NewDCCSend = intDCCSends
End Function

Sub PutText(rtf As RichTextBox, strData As String)
    
    PutData rtf, strData
    Exit Sub
    '* Not Finished
    If strData = "" Then Exit Sub
    'DoEvents
    Dim i As Long, Length As Integer, strChar As String
    Dim strBuffer As String, j As Long, colorTable As String, strelse As String
    Dim strRTF As String, bold As Boolean, underline As Boolean, clr As String
    Dim strRt As String, bclr As Integer, dftclr As Integer
    
    Dim chr As CHARFORMAT2
    chr.cbSize = Len(chr)
    chr.crTextColor = lngForeColor
    chr.crBackColor = lngBackColor
    'chr.bAnimation
    chr.dwMask = CFM_bold Or CFM_COLOR Or CFM_BACKCOLOR Or CFM_FACE Or CFM_UNDERLINETYPE
'    chr
    
    
    
    strData = " " & strData
    Length = Len(strData)
    i = 1
    rtf.SelStart = Len(rtf.Text)
    rtf.SelFontName = strFontName
    
    'DoEvents
    clr = RAnsiColor(lngForeColor)
    bclr = 0
    dftclr = clr
    
    Do
        strChar = Mid(strData, i, 1)
        Select Case strChar
            Case strBold
                DoEvents
                
                strBuffer = ""
                
                bold = Not bold
                'strRt = updateColorTable(CInt(clr), 3, "", rtf.TextRTF, bold, underline, False)
                
                i = i + 1
            Case strUnderline
                'strRt = updateColorTable(CInt(clr), CInt(bclr), strBuffer, CStr(rtf.TextRTF), bold, underline)
                rtf.TextRTF = strRt
                strBuffer = ""
                
                underline = Not underline
                
                'strRt = updateColorTable(CInt(clr), Int(bclr), "", rtf.TextRTF, bold, underline, False)
                
                i = i + 1
            Case strReverse
                i = i + 1
            Case strColor
                'strRt = updateColorTable(CInt(clr), Int(bclr), strBuffer, rtf.TextRTF, bold, underline, False)
                
                strBuffer = ""
                i = i + 1
                If i > Length Then GoTo TheEnd
                Do Until Not ValidColorCode(strBuffer) Or i > Length
                    strBuffer = strBuffer & Mid(strData, i, 1)
                    i = i + 1
                Loop
                If ValidColorCode(strBuffer) And i > Length Then GoTo TheEnd
                strBuffer = LeftR(strBuffer, 1)
                rtf.SelStart = Len(rtf.Text)
                If strBuffer = "" Then
                    rtf.SelColor = lngForeColor
                Else
                    rtf.SelColor = AnsiColor(LeftOf(strBuffer, ","))
                End If
                i = i - 1
                strBuffer = ""

            Case Else
'                MsgBox strBuffer
                strBuffer = strBuffer & strChar
                'strRt = updateColorTable(CInt(clr), CInt(bclr), strBuffer, CStr(rtf.TextRTF))
                'MsgBox strChar & "~~" & vbCrLf & strRt
                i = i + 1
        End Select
        
    Loop Until i >= Length
TheEnd:
    'strRt = updateColorTable(CInt(clr), CInt(bclr), strBuffer, CStr(rtf.TextRTF), bold, underline)
    strBuffer = ""

    'MsgBox rtf.TextRTF
    
    rtf.TextRTF = strRt
    rtf.SelStart = Len(rtf.Text)
    rtf.SelText = vbCrLf
End Sub


Function RAnsiColor(lngColor As Long) As Integer
    Select Case lngColor
        Case RGB(255, 255, 255): RAnsiColor = 0
        Case RGB(0, 0, 0): RAnsiColor = 1
        Case RGB(0, 0, 127): RAnsiColor = 2
        Case RGB(0, 127, 0): RAnsiColor = 3
        Case RGB(255, 0, 0): RAnsiColor = 4
        Case RGB(127, 0, 0): RAnsiColor = 5
        Case RGB(127, 0, 127): RAnsiColor = 6
        Case RGB(255, 127, 0): RAnsiColor = 7
        Case RGB(255, 255, 0): RAnsiColor = 8
        Case RGB(0, 255, 0): RAnsiColor = 9
        Case RGB(0, 148, 144): RAnsiColor = 10
        Case RGB(0, 255, 255): RAnsiColor = 11
        Case RGB(0, 0, 255): RAnsiColor = 12
        Case RGB(255, 0, 255): RAnsiColor = 13
        Case RGB(92, 92, 92): RAnsiColor = 14
        Case RGB(184, 184, 184): RAnsiColor = 15
        Case RGB(0, 0, 0): RAnsiColor = 99
    End Select

End Function

Public Function RichWordOver(rch As RichTextBox, X As Single, Y As Single) As String
    Dim pt As POINTAPI
    Dim pos As Long
    Dim start_pos As Long
    Dim end_pos As Long
    Dim ch As String
    Dim txt As String
    Dim txtlen As Long
    Dim i As Long, j As Long

    ' Convert the position to pixels.
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY

    ' Get the character number
    On Error Resume Next
    pos = SendMessage(rch.hWnd, EM_CHARFROMPOS, 0&, pt)
    
    'Exit Function
    If pos <= 0 Or pos >= Len(rch.Text) Then
        RichWordOver = ""
        Exit Function
    End If
    
    txt = ""
    For i = pos To 1 Step -1
        ch = Mid(rch.Text, i, 1)
        If i Mod 100 = 3 Then DoEvents
        If ch = " " Or _
           ch = "," Or _
           ch = "(" Or _
           ch = ")" Or _
           ch = "]" Or _
           ch = "[" Or _
           ch = "{" Or _
           ch = """" Or _
           ch = "'" Or _
           ch = chr(9) Or _
           ch = "}" Then
            start_pos = i
            GoTo haha
        End If
    Next i
haha:
    txt = ""
    For i = pos To Len(rch.Text)
        ch = Mid(rch.Text, i, 1)
        If ch = " " Or _
           ch = "," Or _
           ch = "(" Or _
           ch = ")" Or _
           ch = "]" Or _
           ch = "[" Or _
           ch = "{" Or _
           ch = "}" Or _
           ch = """" Or _
           ch = "'" Or _
           ch = chr(9) Then
            end_pos = i
            Exit For
        End If
    Next i
    
    If end_pos > Len(rch.Text) Or end_pos <= 0 Then end_pos = Len(rch.Text)
    
    RichWordOver = RightR(Replace(Mid(rch.Text, start_pos, end_pos - start_pos), chr(13), ""), 1)
End Function



Function AnsiColor(intColNum As Integer) As Long
    Select Case intColNum
        Case 0: AnsiColor = RGB(255, 255, 255)
        Case 1: AnsiColor = RGB(0, 0, 0)
        Case 2: AnsiColor = RGB(0, 0, 127)
        Case 3: AnsiColor = RGB(0, 127, 0)
        Case 4: AnsiColor = RGB(255, 0, 0)
        Case 5: AnsiColor = RGB(127, 0, 0)
        Case 6: AnsiColor = RGB(127, 0, 127)
        Case 7: AnsiColor = RGB(255, 127, 0)
        Case 8: AnsiColor = RGB(255, 255, 0)
        Case 9: AnsiColor = RGB(0, 255, 0)
        Case 10: AnsiColor = RGB(0, 148, 144)
        Case 11: AnsiColor = RGB(0, 255, 255)
        Case 12: AnsiColor = RGB(0, 0, 255)
        Case 13: AnsiColor = RGB(255, 0, 255)
        Case 14: AnsiColor = RGB(92, 92, 92)
        Case 15: AnsiColor = RGB(184, 184, 184)
        Case Else: AnsiColor = RGB(0, 0, 0)
    End Select
End Function


Sub ChangeNick(strOldNick As String, strNewNick As String)
    Dim i As Integer, bChangedQuery As Boolean, inttemp As Integer
    
    For i = 1 To intChannels
    
        
        If Channels(i).InChannel(strOldNick) Then
            
            'change in queries :)
            If Not bChangedQuery Then
                inttemp = GetQueryIndex(strOldNick)
                If inttemp <> -1 Then
                    Queries(inttemp).lblNick = strNewNick
                    Queries(inttemp).strNick = strNewNick
                    Queries(inttemp).Caption = strNewNick
                    bChangedQuery = True
                End If
            End If
            
            'change in channel :)
            If Channels(i).strName <> "" Then
                Channels(i).ChangeNck strOldNick, strNewNick
            End If
        End If
    Next i
End Sub

Function Combine(arrItems() As String, intStart As Integer, intEnd As Integer) As String
    '* This returns the given parameter range as one string
    '* -1 specified as strEnd means from intStart to the last parameter
    '* Ex: strParams(1) is "#mIRC", then params 2 thru 16 are the nicks
    '*     of the users in the channel, simply do something like this
    '*     strNames = Params(2, -1)
    '*     that would return all the nicks into one string, similar to mIRC's
    '*      $#- identifier ($2-) in this case
    
    '* Declare variables
    Dim strFinal As String, intLast As Integer, i As Integer
    
    '* check for bad parameters
    If intStart < 1 Or intEnd > UBound(arrItems) + 1 Then Exit Function
    
    '* if intEnd = -1 then set it to param count
    If intEnd = -1 Then intLast = UBound(arrItems) + 1 Else intLast = intEnd
        
    For i = intStart To intLast
        strFinal = strFinal & arrItems(i - 1)
        If i <> intLast Then strFinal = strFinal & " "
    Next i
    
    Combine = Replace(strFinal, vbCr, "")
    strFinal = ""
End Function

Function DisplayNick(nckNick As Nick) As String
    Dim strPre As String
    If nckNick.Voice Then strPre = "+"
    If nckNick.Helper Then strPre = "%"
    If nckNick.Op Then strPre = "@"
    DisplayNick = strPre & nckNick.Nick
End Function

Sub DoMode(strChannel As String, bAdd As Boolean, strMode As String, strParam As String)
    
    If strChannel = strMyNick Then
        If bAdd Then
            Client.AddMode strMode, bAdd
        Else
            Client.RemoveMode strMode
        End If
        Exit Sub
    End If
    
    Dim intX As Integer, i As Integer
    intX = GetChanIndex(strChannel)
    If intX = -1 Then Exit Sub
    
    Select Case strMode
        Case "v"
            Channels(intX).SetVoice strParam, bAdd
            
        Case "o"
            Channels(intX).SetOp strParam, bAdd
            If bAdd Then
                If strParam = strMyNick Then Channels(intX).rtbTopic.Tag = ""
            Else
                If strParam = strMyNick Then Channels(intX).rtbTopic.Tag = "locked"
            End If
        Case "h"
            Channels(intX).SetHelper strParam, bAdd
        Case "b"
        Case "k"
            If bAdd = True Then
                Channels(intX).strKey = strParam
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).strKey = ""
                Channels(intX).RemoveMode strMode
            End If
        Case "l"
            If bAdd = True Then
                Channels(intX).intLimit = CInt(strParam)
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).intLimit = 0
                Channels(intX).RemoveMode strMode
            End If
        Case Else
            If bAdd = True Then
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).RemoveMode strMode
            End If
    End Select
End Sub

Public Function GetAlias(StrChan As String, strData As String) As String
    Dim arrParams() As String, i As Integer, strP As String, strCom As String
    Dim strFinal As String, strAdd As String, bSpace As Boolean, inttemp As Integer
    Dim strTemp As String, strNck As String
    
    DoEvents
    
    If InStr(strData, " ") Then
        Seperate strData, " ", strCom, strData
        arrParams = Split(strData, " ")
    Else
        strCom = strData
        arrParams = Split("", "")
    End If
    bSpace = True
    'DoEvents
    
    For i = LBound(arrParams) To UBound(arrParams)
        strP = arrParams(i)
        strAdd = ""
        If strP = "$+" Then
            strFinal = LeftR(strFinal, 1)
            bSpace = False
        ElseIf left(strP, 1) = "$" Then
            strAdd = GetVar(StrChan, RightR(strP, 1))
        Else
            strAdd = strP
        End If
        
        strFinal = strFinal & strAdd
        If bSpace Then
            strFinal = strFinal & " "
        Else
            bSpace = True
        End If
    Next i
    
    If Len(strFinal) > 0 Then strFinal = LeftR(strFinal, 1)
    
    ReDim arrParams(1) As String
    arrParams = Split(strFinal, " ")
    
    Dim r As String 'return
    Select Case LCase(strCom)
        Case "query", "q"
            strTemp = Combine(arrParams, 2, -1)
            strNck = Combine(arrParams, 1, 1)
            If QueryExists(strNck) Then
                inttemp = GetQueryIndex(strNck)
                Queries(inttemp).PutText strMyNick, strTemp
                r = "PRIVMSG " & strNck & " :" & strTemp
            Else
                inttemp = NewQuery(strNck, "")
                If UBound(arrParams) > 0 Then
                    Queries(inttemp).PutText strMyNick, strTemp
                    r = "PRIVMSG " & strNck & " :" & strTemp
                End If
            End If
                
        Case "msg"
            r = "PRIVMSG " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
            EchoActive strColor & "05[msg]" & strColor & " -> " & strBold & Combine(arrParams, 1, 1) & strBold & ": " & Combine(arrParams, 2, -1)
        Case "me", "action", "describe"
            strTemp = Combine(arrParams, 1, -1)
            r = "PRIVMSG " & StrChan & " :" & strAction & "ACTION " & strTemp & strAction
            If left(StrChan, 1) = "#" Then
                inttemp = GetChanIndex(StrChan)
                If inttemp = -1 Then Exit Function
                PutData Channels(inttemp).DataIn, strColor & "06" & strMyNick & " " & strTemp
            Else
                inttemp = GetQueryIndex(StrChan)
                If inttemp = -1 Then Exit Function
                PutData Queries(inttemp).DataIn, strColor & "06" & strMyNick & " " & strTemp
            End If
        Case "quit", "signoff"
            r = "QUIT :" & Combine(arrParams, 1, -1)
        Case "notice"
            r = "NOTICE " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
            EchoActive strColor & "05[notice]" & strColor & " -> " & strBold & Combine(arrParams, 1, 1) & strBold & ":" & chr(9) & Combine(arrParams, 2, -1)
        Case "raw"
            r = Combine(arrParams, 1, -1)
        Case "nick"
            If Client.sock.State = 0 Then
                strMyNick = Combine(arrParams, 1, 1)
            Else
                r = "NICK " & Combine(arrParams, 1, 1)
            End If
        Case "ea"
            EchoActive Combine(arrParams, 1, 1)
        Case "id"   'identify with nickserv
            r = "PRIVMSG NickServ :IDENTIFY " & Combine(arrParams, 1, 1)
            EchoActive strColor & "05[msg]" & strColor & " -> " & strBold & "NickServ" & strBold & ": IDENTIFY ********"
        Case "part"
            strTemp = Combine(arrParams, 1, -1)
            If UBound(arrParams) = 0 Then
                r = "PART " & StrChan
                strTemp = strTemp
            Else
                r = "PART " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
                strTemp = LeftOf(strTemp, " ")
            End If
            
            inttemp = GetChanIndex(strTemp)
            On Error Resume Next
            Channels(inttemp).Tag = "PARTNOW"
        Case "server"
            strServer = Combine(arrParams, 1, 1)
            If UBound(arrParams) > 0 Then
                strport = Int(Combine(arrParams, 2, 2))
            End If
            If Client.sock.State <> 0 Then Call Client.mnu_File_Disconnect_Click
            TimeOut 0.1
            Call Client.mnu_File_Connect_Click
        Case "sv"
            strVersionReply = Combine(arrParams, 1, -1)
            EchoActive "Version Reply changed to '" & Combine(arrParams, 1, -1) & "'", 6
        Case "join"
            'ok here's how it is, you can type /join #blah (key), #blah2, #blah3, #blah4 (key),
            'so we need special equiptment to handle this
            Dim strChans() As String
            strChans = Split(strData, ",")
            
            For inttemp = LBound(strChans) To UBound(strChans)
                Dim prefix As String
                prefix = ""
                If left(strChans(inttemp), 1) <> "#" And left(strChans(inttemp), 1) <> "&" Then prefix = "#"
                Client.SendData "JOIN " & prefix & strChans(inttemp)
                WriteINI "lag", prefix & strChans(inttemp), GetTickCount()
                TimeOut 0.1
            Next inttemp
        Case "connect"
            If Combine(arrParams, 1, 1) <> "" Then
                strServer = Combine(arrParams, 1, 1)
            End If
            If UBound(arrParams) > 0 Then
                strport = Int(Combine(arrParams, 2, 2))
            End If
            If Client.sock.State <> 0 Then Call Client.mnu_File_Disconnect_Click
            Call Client.mnu_File_Connect_Click
        Case "disconnect"
            Call Client.mnu_File_Disconnect_Click
        Case "bl"
            If BuddyList.Visible Then
                Unload BuddyList
            Else
                Load BuddyList
            End If
        Case "kill"
            r = "KILL " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
        Case "es"
            PutText Status.DataIn, Combine(arrParams, 1, -1)
        Case "list"
            ChannelsList.lvChannels.ListItems.Clear
            If Combine(arrParams, 1, 1) = "" Then
                r = "LIST >0"
            Else
                r = "LIST " & Combine(arrParams, 1, -1)
            End If
        Case "away"
            If Combine(arrParams, 1, -1) = "" Then
                r = "AWAY :"
                If lngGoneFor = 0 Then Else lngGoneFor = CTime - lngGoneFor
                
                '* Announce?
                If bAnnounce Then
                    ActionAll "is back, gone for " & Duration(lngGoneFor)
                End If
                lngGoneFor = 0
            Else
                r = "AWAY :" & Combine(arrParams, 1, -1)
                lngGoneFor = CTime
                
                '* Announce?
                If bAnnounce Then
                    ActionAll "is away, " & Combine(arrParams, 1, -1)
                End If
            End If
        Case "back"
                r = "AWAY :"
                If lngGoneFor = 0 Then Else lngGoneFor = CTime - lngGoneFor
                
                '* Announce?
                If bAnnounce Then
                    ActionAll "is back, gone for " & Duration(lngGoneFor)
                End If
                lngGoneFor = 0
        Case "send", "s"
            If Combine(arrParams, 1, -1) = "" Then Exit Function
            r = "PRIVMSG " & StrChan & " :" & Combine(arrParams, 1, -1)
            If left(StrChan, 1) = "#" Then
                inttemp = GetChanIndex(StrChan)
                Channels(inttemp).PutText strMyNick, Combine(arrParams, 1, -1)
            ElseIf StrChan = "" Then
            Else
                inttemp = GetQueryIndex(StrChan)
                Queries(inttemp).PutText strMyNick, Combine(arrParams, 1, -1)
            End If
        Case "version", "ver"
            r = "PRIVMSG " & Combine(arrParams, 1, 1) & " :" & strAction & "VERSION" & strAction
            EchoActive "[CTCP!" & Combine(arrParams, 1, 1) & "] -> VERSION", 4
        Case "ping", "png"
            r = "PRIVMSG " & Combine(arrParams, 1, 1) & " :" & strAction & "PING " & GetTickCount & strAction
            EchoActive "[CTCP!" & Combine(arrParams, 1, 1) & "] -> PING", 4
        Case "ctcp"
            r = "PRIVMSG " & Combine(arrParams, 1, 1) & " :" & strAction & Combine(arrParams, 2, -1) & strAction
            EchoActive "[CTCP!" & Combine(arrParams, 1, 1) & "] -> " & Combine(arrParams, 2, -1), 4
        Case "kick"
            r = "KICK " & Combine(arrParams, 1, 1) & " " & Combine(arrParams, 2, 2) & " :" & Combine(arrParams, 3, -1)
        Case "k"
            r = "KICK " & StrChan & " " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
        Case "asctime"
            r = ""
            On Error Resume Next
            PutData Status.DataIn, strColor & "10" & strBold & Combine(arrParams, 1, 1) & strBold & " -> " & strBold & AscTime(CLng(Combine(arrParams, 1, 1))) & strBold
            Exit Function
        Case "m"
            r = "MODE " & StrChan & " " & Combine(arrParams, 1, -1)
        Case "clear"
            r = ""
            On Error Resume Next
            Client.ActiveForm.DataIn.Text = ""
            Exit Function
        Case "dccsend"
            Dim strFile As String, dcc As DCC_INFO, lngSize As Long, lngPort As Long, intRet As Integer
            Dim cD As New clsDialog
            If UBound(arrParams) < 1 Then
                On Error Resume Next
                strFile = cD.OpenDialog(Client, "All Files (*.*) |*.*|", "Send " & Combine(arrParams, 1, 1) & " a file", App.path)
            Else
                strFile = Combine(arrParams, 2, -1)
            End If
            'MsgBox strFile
            Randomize
            lngPort = Int(60000 * Rnd) + 4000
            lngSize = FileLen(strFile)
            strFile = RealFile(Replace(strFile, " ", "_"))
            dcc.File = strFile
            dcc.Nick = Combine(arrParams, 1, 1)
            dcc.type = 1    'dcc send
            dcc.Size = lngSize
            dcc.Port = lngPort
            
            intRet = NewDCCSend(dcc)
            If intRet = -1 Then Exit Function
            DCCSends(intRet).sock.Close
            DCCSends(intRet).strFullPath = strFile
            DCCSends(intRet).strNick = Combine(arrParams, 1, 1)
            DCCSends(intRet).sock.LocalPort = lngPort
            DCCSends(intRet).sock.RemotePort = lngPort
            DCCSends(intRet).sock.Listen
            DCCSends(intRet).lngSentRcvd = 0
            DCCSends(intRet).lblStat = "Awaiting acceptance.."
            DCCSends(intRet).lngBps = 0
            DCCSends(intRet).Caption = "DCC Send - 00.00%"
            strFile = LeftOf(strFile, chr(0))
            r = "PRIVMSG " & Combine(arrParams, 1, 1) & " :" & strAction & "DCC SEND " & CStr(strFile) & " " & LongIp(Client.sock.LocalIP) & " " & lngPort & " " & lngSize & strAction
        Case "dccchat"
            If LCase(Combine(arrParams, 1, 1)) = LCase(strMyNick) Then
                EchoActive "* Cannot open DCC Chat session with yourself.", 5
                r = ""
                GoTo haha
            End If
            
            Dim intRet2 As Integer
            Randomize
            lngPort = Int(60000 * Rnd) + 4000
            
            intRet2 = NewDCCChat(Combine(arrParams, 1, 1), "<Unknown>", lngPort)
            If intRet2 = -1 Then
                EchoActive "* Unable to create new DCC Chat window, possible too many currently open", 4
                Exit Function
            End If
            
            DCCChats(intRet2).sock.Close
            DCCChats(intRet2).sock.LocalPort = lngPort
            DCCChats(intRet2).sock.Listen
            PutData DCCChats(intRet2).DataIn, strColor & "02* Requesting DCC Chat session with " & Combine(arrParams, 1, 1)
            
            r = "PRIVMSG " & Combine(arrParams, 1, 1) & " :" & strAction & "DCC CHAT chat " & LongIp(Client.sock.LocalIP) & " " & lngPort & strAction
haha:
        Case "test"
            'MsgBox RAnsiColor(vbRed)
        Case "voice"
            r = "MODE " & StrChan & " +" & String(UBound(arrParams) + 1, "v") & " " & Combine(arrParams, 1, -1)
        Case "op"
            r = "MODE " & StrChan & " +" & String(UBound(arrParams) + 1, "o") & " " & Combine(arrParams, 1, -1)
        Case "halfop", "helper"
            r = "MODE " & StrChan & " +" & String(UBound(arrParams) + 1, "h") & " " & Combine(arrParams, 1, -1)
        Case "dns"
            
            PutData Status.DataIn, strColor & "06 Attempting to resolve " & strBold & Combine(arrParams, 1, 1) & strBold & ", this may cause the client to freeze for a long while."
            Dim retColl As Collection
            Dim nCount As Integer, s As String
            s = " was"
            
            Set retColl = ResolveIpaddress(Combine(arrParams, 1, 1))

            If retColl.Count = 0 Then
                PutData Status.DataIn, vbCrLf & strColor & "06 Unable to resolve " & strBold & Combine(arrParams, 1, 1)
                Exit Function
            End If
            
            If retColl.Count <> 1 Then s = "es were"
            PutData Status.DataIn, ""
            PutData Status.DataIn, strColor & "06" & strBold & retColl.Count & strBold & " ip address" & s & " found for the host name " & strBold & Combine(arrParams, 1, 1)
            
            If retColl.Count > 0 Then
                For nCount = 1 To retColl.Count
                    PutData Status.DataIn, strColor & "06" & "  " & nCount & "." & chr(9) & CStr(retColl.Item(nCount))
                    'lstResolvedAddress.AddItem CStr(retColl.Item(nCount))
                Next nCount
            End If

            Exit Function
        Case "tracert", "trace", "traceroute", "tr"
            PutData Status.DataIn, strColor & "06 Attempting to trace " & strBold & Combine(arrParams, 1, 1) & strBold & ", this may cause the client to freeze for a long while."
            DoEvents
            Dim strHost As String
            strHost = Combine(arrParams, 1, 1)
            vbWSAStartup               ' Initialize Winsock
            
            If Len(strHost) = 0 Then
                strHost = vbGetHostName
            End If
            
            If strHost = "" Then
                PutData Status.DataIn, strColor & "06No hostname specified for tracert function."
                vbWSACleanup
                Exit Function
            End If
            
            strHost = vbGetHostByName(Combine(arrParams, 1, 1))
            vbIcmpCreateFile           ' Get ICMP Handle
           
            
            ' The following determines the TTL of the ICMPEcho for TRACE function
            
            PutData Status.DataIn, strColor & "06 Tracing Route to " & strBold & strHost & strBold & ":" & vbCrLf
            
            For TTL = 2 To 255
                pIPo.TTL = TTL
                vbIcmpSendEcho             ' Send the ICMP Echo Request
                
                RespondingHost = CStr(pIPe.Address(0)) + "." + CStr(pIPe.Address(1)) + "." + CStr(pIPe.Address(2)) + "." + CStr(pIPe.Address(3))
                DoEvents
                If RespondingHost = strHost Then
                    PutData Status.DataIn, strColor & "06 Route Trace has Completed" & vbCrLf & vbCrLf
                    Exit For        ' Stop TraceRT
                End If
            Next TTL
            vbIcmpCloseHandle          ' Close the ICMP Handle
            vbWSACleanup               ' Close Winsock
        Case Else
        
            '****************************************
            '* This is where we implement scripting *
            '****************************************
            strTemp = GetAliasCode(LCase(strCom))
            Client.objScript.Channel = ""
            If strTemp = chr(8) Then    'does NOT exist in aliases
                '* Send raw
                r = strCom & " " & Combine(arrParams, 1, -1)
            Else    'DOES exist!
                Client.objScript.Channel = StrChan
                Client.objScript.MYNick = strMyNick
                CopyParameters arrParams
                
                On Error Resume Next
                Client.cScript.ExecuteStatement strTemp
                
                If Err Then
                    EchoActive "* Occured while trying to executing the '" & strCom & "' alias", 4 ': ERROR #" & Err & " : " & Error, 4
                End If
            End If
    End Select
    
    GetAlias = r
    
End Function


Function SV(lngWhat As Long, strText As String) As String
    If lngWhat = 0 Then
        SV = ""
    ElseIf lngWhat = 1 Then
        SV = lngWhat & " " & strText & " "
    Else
        SV = lngWhat & " " & strText & "s" & " "
    End If
End Function

Sub TimeOut(Duration)
    StartTime = Timer
    Do While Timer - StartTime < Duration
        X = DoEvents()
    Loop
End Sub

Function GetChanIndex(strName As String) As Integer
    Dim i As Integer

    For i = 1 To intChannels
        If LCase(Channels(i).strName) = Replace(LCase(strName), chr(13), "") Then
            GetChanIndex = i
            Exit Function
        End If
    Next i
    GetChanIndex = -1
End Function

Function GetQueryIndex(strNick As String) As Integer
    Dim i As Integer

    For i = 1 To intQueries
        If LCase(Queries(i).strNick) = LCase(strNick) Then
            GetQueryIndex = i
            Exit Function
        End If
    Next i
    GetQueryIndex = -1
End Function

Function GetVar(StrChan As String, strName As String)
    Dim r As String     'r is the return value
    Dim inttemp As String
    
    On Error Resume Next
    Select Case LCase(strName)
        Case "version"
            r = App.Major & "." & App.Minor & App.Revision
        Case "chan", "channel", "ch"
            r = StrChan
        Case "me"
            r = strMyNick
        Case "server"
            r = Client.sock.RemoteHost
        Case "port"
            r = Client.sock.RemotePort
        Case "randnick"
            inttemp = GetChanIndex(StrChan)
            If left(StrChan, "1") = "#" Then
                With Channels(inttemp)
                    Randomize
                    r = .GetNick(Int(Rnd * .intNicks) + 1)
                End With
            End If
        Case "date"
            r = Date
        Case "time"
            r = Time
        Case "usercount", "users", "unum", "ucount", "chancount"
            inttemp = GetChanIndex(StrChan)
            If inttemp = -1 Then
                r = 0
            Else
                r = Channels(inttemp).intUserCount
            End If
        Case "now"
            r = Now
        Case "ctime"
            r = CTime
        Case "ticks"
            r = GetTickCount
        Case "ip"
            r = Client.sock.LocalIP
        Case "host"
            r = Client.sock.LocalHostName
        Case "uptime"
            r = Duration(GetTickCount / 1000)
        Case "info"
            r = "projectIRC " & App.Major & "." & App.Minor & App.Revision & " by vcv ( 12mailto:sappy@adelphia.net )"
    End Select
    GetVar = r
End Function

Function LeftOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strData, strDelim)
    If intPos Then
        LeftOf = left(strData, intPos - 1)
    Else
        LeftOf = strData
    End If
End Function

Function LeftR(strData As String, intMin As Integer)
    
    On Error Resume Next
    LeftR = left(strData, Len(strData) - intMin)
End Function

Function NewChannel(strName As String) As Integer
    Dim i As Integer

    For i = 1 To intChannels
        If Channels(i).strName = "" Then
            Channels(i).strName = strName
            Channels(i).Visible = True
            Channels(i).lblName = strName
            Channels(i).Tag = i
            Channels(i).Update
            NewChannel = i
            Exit Function
        End If
    Next i
    intChannels = intChannels + 1
    Set Channels(intChannels) = New Channel
    Channels(intChannels).strName = strName
    Channels(intChannels).Visible = True
    Channels(intChannels).Update
    Channels(intChannels).lblName = strName
    Channels(intChannels).Caption = strName
    
    Channels(intChannels).Tag = intChannels
    NewChannel = intChannels
End Function
Function NewQuery(strNick As String, strHost As String) As Integer
    Dim i As Integer, strHostX As String
    strHostX = RightOf(strHost, "!")
    
    i = GetQueryIndex(strNick)
    If i <> -1 Then
        Queries(i).SetFocus
        Exit Function
    End If

    For i = 1 To intQueries
        If Queries(i).strNick = "" Then
            Queries(i).Caption = strNick
            Queries(i).lblNick = strNick
            Queries(i).strNick = strNick
            Queries(i).strHost = strHostX
            Queries(i).lblHost = strHostX
            Queries(i).Visible = True
            Queries(i).Tag = i
            NewQuery = i
            Exit Function
        End If
    Next i
    
    intQueries = intQueries + 1
    Set Queries(intQueries) = New Query
    Queries(intQueries).strNick = strNick
    Queries(intQueries).lblNick = strNick
    Queries(intQueries).Caption = strNick
    Queries(intQueries).lblHost = strHostX
    Queries(intQueries).strHost = strHostX
    Queries(intQueries).Visible = True
    Queries(intQueries).Tag = intQueries
    NewQuery = intQueries
End Function

Sub NickQuit(strNick As String, strMsg As String)
    For i = 1 To intChannels
        If Channels(i).InChannel(strNick) And Channels(i).strName <> "" Then
            Channels(i).RemoveNick strNick
            PutData Channels(i).DataIn, strColor & "02" & strBold & strNick & strBold & " has Quit IRC [ " & strMsg & " ]"
        End If
    Next i

    For i = 1 To intQueries
        If LCase(Queries(i).strNick) = LCase(strNick) Then
            PutData Queries(i).DataIn, strColor & "02" & strBold & strNick & strBold & " has Quit IRC [ " & strMsg & " ]"
            Exit Sub
        End If
    Next i
End Sub

Function params(parsed As ParsedData, intStart As Integer, intEnd As Integer) As String
    '* This returns the given parameter range as one string
    '* -1 specified as strEnd means from intStart to the last parameter
    '* Ex: strParams(1) is "#mIRC", then params 2 thru 16 are the nicks
    '*     of the users in the channel, simply do something like this
    '*     strNames = Params(2, -1)
    '*     that would return all the nicks into one string, similar to mIRC's
    '*      $#- identifier ($2-) in this case
    
    '* Declare variables
    Dim strFinal As String, intLast As Integer, i As Integer
    
    '* check for bad parameters
    If intStart < 1 Or intEnd > parsed.intParams Then Exit Function
    
    '* if intEnd = -1 then set it to param count
    If intEnd = -1 Then intLast = parsed.intParams Else intLast = intEnd
        
    For i = intStart To intLast
        strFinal = strFinal & parsed.strParams(i)
        If i <> intLast Then strFinal = strFinal & " "
    Next i
    
    params = Replace(strFinal, vbCr, "")
    strFinal = ""
End Function

Sub ParseData(ByVal strData As String, ByRef parsed As ParsedData)

    '* Declare variables
    Dim strTMP As String, i As Integer
    
    '* Reset variables
    bHasPrefix = False
    parsed.strNick = ""
    parsed.strIdent = ""
    parsed.strHost = ""
    parsed.strCommand = ""
    parsed.intParams = 1
    ReDim parsed.strParams(1 To 1) As String
    
    '* Check for prefix, if so, parse nick, ident and host (or just host)
    If left(strData, 1) = ":" Then
        bHasPrefix = True
        strData = Right(strData, Len(strData) - 1)
        '* Put data left of " " in strHost, data right of " "
        '* into strData
        Seperate strData, " ", parsed.strHost, strData
        parsed.strFullHost = parsed.strHost
        
        '* Check to see if client host name
        If InStr(parsed.strHost, "!") Then
            Seperate parsed.strHost, "!", parsed.strNick, parsed.strHost
            Seperate parsed.strHost, "@", parsed.strIdent, parsed.strHost
        End If
    End If
    
    '* If any params, parse
    If InStr(strData, " ") Then
        Seperate strData, " ", parsed.strCommand, strData
        
        parsed.AllParams = strData
       '* Let's parse all the parameters.. yummy
Begin: '* OH NO I USED A LABEL!

        '* If begginning of param is :, indicates that its the last param
        If left(strData, 1) = ":" Then
            parsed.strParams(parsed.intParams) = Right(strData, Len(strData) - 1)
            GoTo Finish
        End If
        '* If there is a space still, there is more params
        If InStr(strData, " ") Then
            Seperate strData, " ", parsed.strParams(parsed.intParams), strData
            parsed.intParams = parsed.intParams + 1
            ReDim Preserve parsed.strParams(1 To parsed.intParams) As String
            GoTo Begin
        Else
            parsed.strParams(parsed.intParams) = strData
        End If
    Else
        '* No params, strictly command
        parsed.intParams = 0
        parsed.strCommand = strData
    End If
Finish:
End Sub

Sub ParseMode(strChannel As String, strData As String)
    Dim strModes() As String, strChar As String
    Dim i As Integer, intParam As Integer
    Dim bAdd As Boolean
    
    bAdd = True
    strModes = Split(strData, " ")
    For i = 1 To Len(strModes(0))
        strChar = Mid(strModes(0), i, 1)
        If left(strChannel, 1) = "#" Then
            Select Case strChar
                Case "+"
                    bAdd = True
                Case "-"
                    bAdd = False
                Case "v", "b", "o", "h", "k", "l", "q", "a"
                    intParam = intParam + 1
                    DoMode strChannel, bAdd, strChar, strModes(intParam)
                Case Else
                    DoMode strChannel, bAdd, strChar, ""
            End Select
        Else    'server
            Select Case strChar
                Case "+"
                    bAdd = True
                Case "-"
                    bAdd = False
                Case "" 'ignore this
                    intParam = intParam + 1
                    DoMode strChannel, bAdd, strChar, strModes(intParam)
                Case Else
                    DoMode strChannel, bAdd, strChar, ""
            End Select

        End If
    Next i
End Sub

Sub PutData(rtf As RichTextBox, strData As String)
    
    If InStr(strData, strBold) Or _
       InStr(strData, strUnderline) Or _
       InStr(strData, strReverse) Or _
       InStr(strData, strColor) Then
    Else
       rtf.SelStart = Len(rtf.Text)
       rtf.SelColor = lngForeColor
       rtf.SelBold = False
       rtf.SelUnderline = False
       rtf.SelStrikeThru = False
       rtf.SelFontName = strFontName
       rtf.SelFontSize = intFontSize
       rtf.SelText = " " & strData & vbCrLf
       Exit Sub
    End If
    
    If strData = "" Then Exit Sub
    'DoEvents
    Dim i As Long, Length As Integer, strChar As String, strBuffer As String
    Dim clr As Integer, bclr As Integer, dftclr As Integer
    
    dftclr = RAnsiColor(lngForeColor)
    strData = " " & strData
    Length = Len(strData)
    i = 1
    rtf.SelStart = Len(rtf.Text)
    rtf.SelColor = lngForeColor
    rtf.SelBold = False
    rtf.SelUnderline = False
    rtf.SelStrikeThru = False
    rtf.SelFontName = strFontName
    rtf.SelFontSize = intFontSize
    
    Do
        strChar = Mid(strData, i, 1)
        Select Case strChar
            Case strBold, chr(15)
                rtf.SelStart = Len(rtf.Text)
                rtf.SelText = strBuffer
                strBuffer = ""
                rtf.SelBold = Not rtf.SelBold
                i = i + 1
            Case strUnderline
                rtf.SelStart = Len(rtf.Text)
                rtf.SelText = strBuffer
                strBuffer = ""
                rtf.SelUnderline = Not rtf.SelUnderline
                i = i + 1
            Case strReverse
                rtf.SelStart = Len(rtf.Text)
                rtf.SelText = strBuffer
                strBuffer = ""
                rtf.SelStrikeThru = Not rtf.SelStrikeThru
                i = i + 1
            Case strColor
                rtf.SelStart = Len(rtf.Text)
                rtf.SelText = strBuffer
                strBuffer = ""
                i = i + 1

                Do Until Not ValidColorCode(strBuffer) Or i > Length
                    strBuffer = strBuffer & Mid(strData, i, 1)
                    i = i + 1
                Loop
                
                strBuffer = LeftR(strBuffer, 1)
                rtf.SelStart = Len(rtf.Text)
                
                
                If strBuffer = "" Then
                    rtf.SelColor = lngForeColor
                Else
                    rtf.SelColor = AnsiColor(LeftOf(strBuffer, ","))
                End If
                
                i = i - 1
                If i >= Length Then GoTo TheEnd
                strBuffer = ""
            Case Else
                strBuffer = strBuffer & strChar
                i = i + 1
        End Select
    Loop Until i > Length
    If strBuffer <> "" Then
            rtf.SelStart = Len(rtf.Text)
            rtf.SelText = strBuffer
            strBuffer = ""
    End If
TheEnd:
    rtf.SelBold = False
    rtf.SelUnderline = False
    rtf.SelStrikeThru = False
    rtf.SelStart = Len(rtf.Text)
    rtf.SelText = vbCrLf
    
End Sub
Function QueryExists(strNick As String) As Boolean
    Dim i As Integer

    For i = 1 To intQueries
        If LCase(Queries(i).strNick) = LCase(strNick) Then
            QueryExists = True
            Exit Function
        End If
    Next i
    QueryExists = False
End Function

Function RealNick(strNick As String) As String
    strNick = Replace(strNick, "@", "")
    strNick = Replace(strNick, "%", "")
    strNick = Replace(strNick, "+", "")
    RealNick = strNick
End Function

Sub RefreshList(lstBox As ListBox)
    'lstBox.AddItem "", 0
    'lstBox.RemoveItem 0
End Sub

Function RightOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        RightOf = Mid(strData, intPos + 1, Len(strData) - intPos)
    Else
        RightOf = strData
    End If
End Function


Function RightR(strData As String, intMin As Integer)
    On Error Resume Next
    RightR = Right(strData, Len(strData) - intMin)
End Function

Sub Seperate(strData As String, strDelim As String, ByRef strLeft As String, ByRef strRight As String)
    '* Seperates strData into 2 variables based on strDelim
    '* Ex: strData is "Bill Clinton"
    '*     Dim strFirstName As String, strLastName As String
    '*     Seperate strData, " ", strFirstName, strLastName
    
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        strLeft = left(strData, intPos - 1)
        strRight = Mid(strData, intPos + 1, Len(strData) - intPos)
    Else
        strLeft = strData
        strRight = strData
    End If
End Sub


Function ValidColorCode(strCode As String) As Boolean
    'MsgBox strCode
    Dim c1 As Integer, c2 As Integer
    If strCode Like "" Or _
       strCode Like "#" Or _
       strCode Like "##" Or _
       strCode Like "#,#" Or _
       strCode Like "##,#" Or _
       strCode Like "#,##" Or _
       strCode Like "#," Or _
       strCode Like "##," Or _
       strCode Like "##,##" Or _
       strCode Like ",#" Or _
       strCode Like ",##" Then
        Dim strCol() As String
        strCol = Split(strCode, ",")
        'DoEvents
        If UBound(strCol) = -1 Then
            ValidColorCode = True
        ElseIf UBound(strCol) = 0 Then
            If strCol(0) = "" Then strCol(0) = 0
            'MsgBox Int(strCol(0)) & " before is ?"
            If Int(strCol(0)) >= 0 And Int(strCol(0)) <= 99 Then
                'MsgBox Int(strCol(0)) & " after is true"
                ValidColorCode = True
                Exit Function
            Else
                ValidColorCode = False
                Exit Function
            End If
        Else
            If strCol(0) = "" Then strCol(0) = lngForeColor
            If strCol(1) = "" Then strCol(1) = 0
            c1 = Int(strCol(0))
            c2 = Int(strCol(1))
            If Int(c2) < 0 Or Int(c2) > 99 Then
                ValidColorCode = False
                Exit Function
            Else
                ValidColorCode = True
                Exit Function
            End If
        End If
        ValidColorCode = True
        Exit Function
    Else
        ValidColorCode = False
        Exit Function
    End If
End Function


