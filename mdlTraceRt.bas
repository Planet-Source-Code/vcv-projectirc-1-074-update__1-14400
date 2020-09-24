Attribute VB_Name = "TraceRt"
'* Original author unknown

' WSock32 Variables

Public iReturn As Long, sLowByte As String, sHighByte As String
Public sMsg As String, HostLen As Long, Host As String
Dim Hostent As Hostent, PointerToPointer As Long, ListAddress As Long
Dim WSAdata As WSAdata, DotA As Long, DotAddr As String, ListAddr As Long
Public MaxUDP As Long, MaxSockets As Long, i As Integer
Public Description As String, Status As String

' ICMP Variables

Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim Addr As Long
Dim RCode As String
Public RespondingHost As String

' TRACERT Variables

Dim TraceRt As Boolean
Dim TTL As Integer


' WSock32 Constants

Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0
Sub vbIcmpCloseHandle()
  
    bReturn = IcmpCloseHandle(hIP)
    
    If bReturn = False Then
        MsgBox "ICMP Closed with Error", vbOKOnly, "VB4032-ICMPEcho"
    End If

End Sub

Sub GetRCode()

    If pIPe.Status = 0 Then RCode = "Success"
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Dest Network Not Reachable"
    If pIPe.Status = 11003 Then RCode = "Dest Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Dest Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Dest Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Rqst Timed Out"
    If pIPe.Status = 11011 Then RCode = "Bad Request"
    If pIPe.Status = 11012 Then RCode = "Bad Route"
    If pIPe.Status = 11013 Then RCode = "TTL Exprd in Transit"
    If pIPe.Status = 11014 Then RCode = "TTL Exprd Reassemb"
    If pIPe.Status = 11015 Then RCode = "Parameter Problem"
    If pIPe.Status = 11016 Then RCode = "Source Quench"
    If pIPe.Status = 11017 Then RCode = "Option too Big"
    If pIPe.Status = 11018 Then RCode = " Bad Destination"
    If pIPe.Status = 11019 Then RCode = "Address Deleted"
    If pIPe.Status = 11020 Then RCode = "Spec MTU Change"
    If pIPe.Status = 11021 Then RCode = "MTU Change"
    If pIPe.Status = 11022 Then RCode = "Unload"
    If pIPe.Status = 11050 Then RCode = "General Failure"
    RCode = RCode + " (" + CStr(pIPe.Status) + ")"

    DoEvents
           
    EchoActive strColor & "06 -> Hop #" & strBold & CStr(pIPo.TTL - 1) & strBold & " : [" & RespondingHost & "] Latency = " & strBold & Trim$(CStr(pIPe.RoundTripTime)) & strBold & "ms TTL = " & strBold & Trim$(CStr(pIPe.Options.TTL)) & strBold
    'EchoActive strColor & "06 -> Hop #" & strBold & CStr(pIPo.TTL - 1) & strBold & " " & RespondingHost & vbCrLf

End Sub
Function HiByte(ByVal wParam As Integer)

    HiByte = wParam \ &H100 And &HFF&

End Function
Function LoByte(ByVal wParam As Integer)

    LoByte = wParam And &HFF&

End Function
Function vbGetHostByName(strHost As String)

    Dim szString As String

    Host = Trim$(strHost)               ' Set Variable Host to Value in Text1.text

    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then              ' If WSock32 error, then tell me about it
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        vbGetHostByName = ""
    Else
        PointerToPointer = gethostbyname(Host)              ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent.h_name, ByVal _
        PointerToPointer, Len(Hostent)                      ' Copy Winsock structure to the VisualBasic structure

        ListAddress = Hostent.h_addr_list                   ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4           ' Copy Winsock structure to the VisualBasic structure
        CopyMemory IPLong, ByVal ListAddr, 4                ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4

        vbGetHostByName = Trim$(CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) _
            + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1)))
    End If

End Function
Function vbGetHostName()
  
    Host = String(64, &H0)          ' Set Host value to a bunch of spaces
    
    If gethostname(Host, HostLen) = SOCKET_ERROR Then     ' This routine is where we get the host's name
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())  ' If WSOCK32 error, then tell me about it
        vbGetHostName = ""
    Else
        Host = left$(Trim$(Host), Len(Trim$(Host)) - 1)   ' Trim up the results
        vbGetHostName = Host                                 ' Display the host's name in label1
    End If

End Function
Sub vbIcmpCreateFile()

    hIP = IcmpCreateFile()

    If hIP = 0 Then
        MsgBox "Unable to Create File Handle", vbOKOnly, "VBPing32"
    End If

End Sub
Sub vbIcmpSendEcho()

    Dim NbrOfPkts As Integer, numCPP As Integer, numPck As Integer
    numCPP = 32
    numPck = 1

    szBuffer = "abcdefghijklmnopqrstuvwabcdefghijklmnopqrstuvwabcdefghijklmnopqrstuvwabcdefghijklmnopqrstuvwabcdefghijklmnopqrstuvwabcdefghijklm"

    If Val(numCPP) < 32 Then numCPP = 32
    If Val(numCPP) > 128 Then numCPP = 128

    szBuffer = left$(szBuffer, CLng(numCPP))

    If IsNumeric(numPck) Then
        If numPck < 1 Then numPck = 1
    Else
        numPck = 1
    End If

    'If TraceRt = True Then Text4.Text = "1"

    For NbrOfPkts = 1 To numPck

        DoEvents
        bReturn = IcmpSendEcho(hIP, Addr, szBuffer, Len(szBuffer), pIPo, pIPe, Len(pIPe) + 8, 2700)

        If bReturn Then

            RespondingHost = CStr(pIPe.Address(0)) + "." + CStr(pIPe.Address(1)) + "." + CStr(pIPe.Address(2)) + "." + CStr(pIPe.Address(3))

            GetRCode

        Else        ' I hate it when this happens.  If I get an ICMP timeout
                    ' during a TRACERT, try again.

            If TraceRt Then
                TTL = TTL - 1
            Else    ' Don't worry about trying again on a PING, just timeout
                'RespondingHost = CStr(pIPe.Address(0)) + "." + CStr(pIPe.Address(1)) + "." + CStr(pIPe.Address(2)) + "." + CStr(pIPe.Address(3))
                EchoActive strColor & "06 - ICMP Request Timeout on " & strBold & RespondingHost & vbCrLf
            End If

        End If

    Next NbrOfPkts

End Sub

Sub vbWSACleanup()

    ' Subroutine to perform WSACleanup

    iReturn = WSACleanup()

    If iReturn <> 0 Then       ' If WSock32 error, then tell me about it.
        sMsg = "WSock32 Error - " & Trim$(Str$(iReturn)) & " occurred in Cleanup"
        MsgBox sMsg, vbOKOnly, "VB4032-ICMPEcho"
        End
    End If

End Sub
Sub vbWSAStartup()
    
    ' Subroutine to Initialize WSock32

    iReturn = WSAStartup(&H101, WSAdata)

    If iReturn <> 0 Then    ' If WSock32 error, then tell me about it
        MsgBox "WSock32.dll is not responding!", vbOKOnly, "VB4032-ICMPEcho"
    End If

    If LoByte(WSAdata.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAdata.wVersion) = WS_VERSION_MAJOR And HiByte(WSAdata.wVersion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(WSAdata.wVersion)))
        sLowByte = Trim$(Str$(LoByte(WSAdata.wVersion)))
        
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported "
        MsgBox sMsg, vbOKOnly, "VB4032-ICMPEcho"
        End
    End If

    If WSAdata.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg, vbOKOnly, "VB4032-ICMPEcho"
        End
    End If
    
    MaxSockets = WSAdata.iMaxSockets

    '  WSAdata.iMaxSockets is an unsigned short, so we have to convert it to a signed long

    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If

    MaxUDP = WSAdata.iMaxUdpDg
    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If

    '  Process the Winsock Description information
 
    Description = ""

    For i = 0 To WSADESCRIPTION_LEN
        If WSAdata.szDescription(i) = 0 Then Exit For
        Description = Description + chr$(WSAdata.szDescription(i))
    Next i

    '  Process the Winsock Status information

    zStatus = ""

    For i = 0 To WSASYS_STATUS_LEN
        If WSAdata.szSystemStatus(i) = 0 Then Exit For
        zStatus = zStatus + chr$(WSAdata.szSystemStatus(i))
    Next i

End Sub

