VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Channel      As String
Public Nick         As String
Public MYNick       As String
Public Active       As String

Private strOut As String

Function Dur(lngSeconds As Long) As String
    If lngSeconds = 0 Then Dur = "0 seconds": Exit Function
    Dim yrs As Long, mnths As Long, dys As Long, hrs As Long, mns As Long, scs As Long
    
    yrs = Int(lngSeconds / 31557600) '# of seconds in a year
    lngSeconds = lngSeconds Mod 31557600
    
    mnths = lngSeconds / 2629800    '# of seconds in month
    lngSeconds = lngSeconds Mod 2629800
    
    dys = lngSeconds / 86400    '# of seconds in a day
    lngSeconds = lngSeconds Mod 86400
    
    hrs = lngSeconds / 3600     '# of seconds in an hour
    lngSeconds = lngSeconds Mod 3600
    
    mns = Int(lngSeconds / 60)
    scs = lngSeconds Mod 60
    
    Dur = LeftR(SV(yrs, "year") & _
               SV(mnths, "month") & _
               SV(dys, "day") & _
               SV(hrs, "hour") & _
               SV(mns, "minute") & _
               SV(scs, "second"), 1)
    
               
End Function


Public Sub ExecAlias(strAlias As String, strData As String)
    Dim strCode As String
    
    strData = GetAliasCode(strAlias)
    If strData = chr(8) Then Exit Sub
    Client.cScript.ExecuteStatement strData
End Sub

Public Function AscTime(lngTime As Long)
'    AscTime = AscTime(lngTime)
End Function


Public Sub Away(Optional strMsg As String = "")
    strOut = GetAlias(Channel, "AWAY " & strMsg)
End Sub

Public Sub Back()
    strOut = GetAlias(Channel, "back")
    Client.SendData strOut
End Sub


Public Function ChanCount(strChannel As String) As Integer
    Dim inttemp As Integer
    inttemp = GetChanIndex(strChannel)
    
    If inttemp = -1 Then
        ChanCount = 0
    Else
        ChanCount = Channels(inttemp).intUserCount
    End If
    
End Function

Public Sub Connect(strServer As String, Optional lngPort As Long = 6667)
    strOut = GetAlias(Channel, "connect " & strServer & " " & lngPort)
End Sub

Public Sub CTCP(strWho As String, strWhat As String)
    strOut = GetAlias(Channel, "ctcp " & strWho & " " & strWhat)
    Client.SendData strOut
End Sub

Public Sub DCCChat(strWho As String)
    strOut = GetAlias(Channel, "dccchat " & strWho)
    Client.SendData strOut
End Sub

Public Sub DCCSend(strWho As String, Optional strFile As String = "")
    strOut = GetAlias(Channel, "dccsend " & strWho & " " & strFile)
    Client.SendData strOut
End Sub
Public Sub describe(strText As String, Optional strChannel As String = "")
    If strChannel = "" Then strChannel = Channel
    strOut = GetAlias(strChannel, "ME " & strText)
    Client.SendData strOut
End Sub


Public Sub Disconnect()
    strOut = GetAlias(Channel, "disconnect")
End Sub


Public Sub ECHO(strCaption As String, strText As String)
    
    Dim strTemp As String, inttemp As Integer
    
    If left(strCaption, 1) = "#" Then
        inttemp = GetChanIndex(strCaption)
        If inttemp = -1 Then Exit Sub
        PutData Channels(inttemp).DataIn, strText
    ElseIf LCase(strCaption) = "status" Then
        PutData Status.DataIn, strText
    ElseIf InStr(strCaption, "DCC Chat - ") Then
        PutData DCCChats(GetDCCChatIndex(RightOf(strCaption, " - "))).DataIn, strText
    ElseIf strCaption = "@" Then
        On Error Resume Next
        PutData Client.ActiveForm.DataIn, strText
        
        '* If not active window has datain, echo to status
        If Err Then PutData Status.DataIn, strText
    Else
        inttemp = GetQueryIndex(strCaption)
        If inttemp = -1 Then Exit Sub
        PutData Queries(inttemp).DataIn, strText
    End If
        
End Sub



Function args(intStart As Integer, Optional intEnd As Integer = -1) As String
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
    If intStart < 1 Or intEnd > UBound(parameters) + 1 Then Exit Function
    
    '* if intEnd = -1 then set it to param count
    If intEnd = -1 Then intLast = UBound(parameters) + 1 Else intLast = intEnd
        
    intStart = intStart + 1
    For i = intStart To intLast
        strFinal = strFinal & parameters(i - 1)
        If i <> intLast Then strFinal = strFinal & " "
    Next i
    
    args = Replace(strFinal, vbCr, "")
    strFinal = ""
End Function

Function arg(intWhich As Integer) As String
    '* This returns the given parameter range as one string
    '* -1 specified as strEnd means from intStart to the last parameter
    '* Ex: strParams(1) is "#mIRC", then params 2 thru 16 are the nicks
    '*     of the users in the channel, simply do something like this
    '*     strNames = Params(2, -1)
    '*     that would return all the nicks into one string, similar to mIRC's
    '*      $#- identifier ($2-) in this case
    
    '* Declare variables
    Dim strFinal As String
        
    If intWhich > UBound(parameters) Then intWhich = UBound(parameters)
    strFinal = parameters(intWhich)
    
    arg = Replace(strFinal, vbCr, "")
    strFinal = ""
End Function

Public Sub Id(strPw As String)
    strOut = GetAlias(Channel, "ID " & strPw)
    Client.SendData strOut
End Sub



Public Sub Kick(strWho As String, strMsg As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    strOut = GetAlias(StrChan, "kick " & strWho & " " & StrChan & " " & strMsg)
    Client.SendData strOut
End Sub
Public Sub Kill(strUser As String, Optional strMsg As String = "Bye")
    strOut = GetAlias(Channel, "kill " & strUser & " " & strMsg)
    Client.SendData strOut
End Sub

Public Sub mode(strMode As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    strOut = GetAlias(StrChan, "mode " & StrChan & " " & strMode)
    Client.SendData strOut
End Sub
Function IsOp(strNick As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    Dim inttemp As Integer
    inttemp = GetChanIndex(StrChan)
    If intttemp = -1 Then IsOp = False: Exit Function
    
    If Channels(inttemp).IsOp(strNick) Then
        IsOp = True
    Else
        IsOp = False
    End If
    
End Function
Function IsHelper(strNick As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    Dim inttemp As Integer
    inttemp = GetChanIndex(StrChan)
    If intttemp = -1 Then IsHelper = False: Exit Function
    
    If Channels(inttemp).isHalfOp(strNick) Then
        IsHelper = True
    Else
        IsHelper = False
    End If
    
End Function
Function IsReg(strNick As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    Dim inttemp As Integer
    inttemp = GetChanIndex(StrChan)
    If intttemp = -1 Then IsReg = False: Exit Function
    
    If Channels(inttemp).isHalfOp(strNick) Then
        IsReg = False
    ElseIf Channels(inttemp).IsOp(strNick) Then
        IsReg = False
    ElseIf Channels(inttemp).IsVoice(strNick) Then
        IsReg = False
    Else
        IsReg = True
    End If
    
End Function

Function IsVoice(strNick As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    Dim inttemp As Integer
    inttemp = GetChanIndex(StrChan)
    If intttemp = -1 Then IsVoice = False: Exit Function
    
    If Channels(inttemp).IsVoice(strNick) Then
        IsVoice = True
    Else
        IsVoice = False
    End If
    
End Function

Public Sub msg(strNick As String, strMsg As String)
    Dim strSend As String
    strSend = GetAlias(Channel, "MSG " & strNick & " " & strMsg)
    Client.SendData strSend
End Sub


Public Sub CJoin(strChn As String)
    Dim strSnd As String
    strSnd = GetAlias(Channel, "JOIN " & strChn)
End Sub
Public Function GetNick(StrChan As String, intIndex As Integer) As String
    Dim inttemp As Integer
    inttemp = GetChanIndex(StrChan)
    
    If inttemp = -1 Then
        GetNick = ""
    Else
        If intWhich > Channels(inttemp).lstNicks.ListCount Then
            GetNick = ""
        Else
            GetNick = RealNick(Channels(inttemp).lstNicks.List(intIndex + 1))
        End If
    End If
    
End Function

Public Sub Notice(strNick As String, strMsg As String)
    strOut = GetAlias(Channel, "NOTICE " & strNick & " " & strMsg)
    Client.SendData strOut
End Sub

Public Sub Part(Optional StrChan As String = "", Optional strMsg As String = "")
    Dim strSend As String
    If StrChan = "" Then StrChan = Channel
    strSend = GetAlias(Channel, "PART " & StrChan & " " & strMsg)
    Client.SendData strSend
End Sub

Public Sub pause(dblLength As Double)
    TimeOut dblLength
End Sub


Public Sub Ping(strUser As String)
    strOut = GetAlias(Channel, "PING " & strUser)
    Client.SendData strOut
End Sub


Public Sub Query(strNick As String, Optional strText As String = "")
    Dim strSend As String
    
    strSend = GetAlias(Channel, "QUERY " & strNick & " " & strText)
    Client.SendData strSend
End Sub


Public Sub Quit(Optional strMsg As String = "using projectIRC, closed")
    strOut = GetAlias(Channel, "QUIT " & strMsg)
    Client.SendData strOut
End Sub


Public Sub Raw(strData As String)
    Client.SendData strData
End Sub


Public Sub Server(strServer As String, Optional lngPort As Long = 6667)
    strOut = GetAlias(Channel, "SERVER " & strServer & " " & lngPort)
End Sub



Public Sub ShellEx(What As Variant)
    'as much as i hate variants, this prevents the need to type-cast within the scripting
    'when passing parameters
    ShellExecute 0, "open", What, "", "", 0&
End Sub

Function uptime() As String
    uptime = Dur(GetTickCount() / 1000)
End Function

Public Sub Version(strUser As String)
    strOut = GetAlias(Channel, "VERSION " & strUser)
    Client.SendData strOut
End Sub



Public Sub Voice(strWho As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    strOut = GetAlias(StrChan, "voice " & strWho)
    Client.SendData strOut
End Sub
Public Sub Op(strWho As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    strOut = GetAlias(StrChan, "op " & strWho)
    Client.SendData strOut
End Sub
Public Sub Helper(strWho As String, Optional StrChan As String = "")
    If StrChan = "" Then StrChan = Channel
    strOut = GetAlias(StrChan, "helper " & strWho)
    Client.SendData strOut
End Sub
