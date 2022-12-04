Attribute VB_Name = "basPing"
'* Original author unknown

Option Explicit

Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)
Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)
Private Const PING_TIMEOUT = 200
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1

Private Const AF_UNSPEC As Integer = 0                    ' unspecified
Private Const AF_UNIX As Integer = 1                      ' local to host (pipes, portals)
Private Const AF_INET As Integer = 2                     ' internetwork: UDP, TCP, etc.
Private Const AF_IMPLINK As Integer = 3                  ' arpanet imp addresses
Private Const AF_PUP As Integer = 4                      ' pup protocols: e.g. BSP
Private Const AF_CHAOS As Integer = 5                    ' mit CHAOS protocols
Private Const AF_IPX As Integer = 6                      ' IPX and SPX
Private Const AF_NS As Integer = AF_IPX                  ' XEROX NS protocols
Private Const AF_ISO As Integer = 7                      ' ISO protocols
Private Const AF_OSI As Integer = AF_ISO                 ' OSI is ISO
Private Const AF_ECMA As Integer = 8                     ' european computer manufacturers
Private Const AF_DATAKIT As Integer = 9                  ' datakit protocols
Private Const AF_CCITT As Integer = 10                    ' CCITT protocols, X.25 etc
Private Const AF_SNA As Integer = 11                      ' IBM SNA
Private Const AF_DECnet As Integer = 12                   ' DECnet
Private Const AF_DLI As Integer = 13                      ' Direct data link interface
Private Const AF_LAT As Integer = 14                      ' LAT
Private Const AF_HYLINK As Integer = 15                  ' NSC Hyperchannel
Private Const AF_APPLETALK As Integer = 16               ' AppleTalk
Private Const AF_NETBIOS As Integer = 17                  ' NetBios-style addresses
Private Const AF_VOICEVIEW As Integer = 18               ' VoiceView
Private Const AF_FIREFOX As Integer = 19                  ' Protocols from Firefox
Private Const AF_UNKNOWN1 As Integer = 20                 ' Somebody is using this!
Private Const AF_BAN As Integer = 21                     ' Banyan
Private Const AF_ATM As Integer = 22                     ' Native ATM Services
Private Const AF_INET6 As Integer = 23                   ' Internetwork Version 6
Private Const AF_CLUSTER As Integer = 24                 ' Microsoft Wolfpack
Private Const AF_12844 As Integer = 25                   ' IEEE 1284.4 WG AF

Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128

Private Type Inet_address
  Byte4 As Byte
  Byte3 As Byte
  Byte2 As Byte
  Byte1 As Byte
End Type
Private IPLong As Inet_address


Private Type ICMP_OPTIONS
    TTL             As Byte
    Tos             As Byte
    flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Private Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long  'formerly integer
  '  Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    data            As String * 250
End Type

Private Type Hostent
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Long
    wMaxUDPDG As Long
    dwVendorInfo As Long
End Type

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (Addr As Long, addrLen As Long, addrType As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal ipaddress$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)



'Public pIPo As IP_OPTION_INFORMATION



'Public pIPe As IP_ECHO_REPLY

'Winsock
Declare Function gethostbyname& Lib "wsock32.dll" (ByVal hostname$)

'Kernel
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Function GetStatusCode(Status As Long) As String

   Dim msg As String

   Select Case Status
      Case IP_SUCCESS:               msg = "ip success"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   
   GetStatusCode = CStr(Status) & "   [ " & msg & " ]"
   
End Function


Private Function HiByte(ByVal wParam As Long) As Integer

    HiByte = wParam \ &H100 And &HFF&

End Function


Private Function LoByte(ByVal wParam As Long) As Integer

    LoByte = wParam And &HFF&

End Function


Private Function PingAddress(szAddress As String, ECHO As ICMP_ECHO_REPLY, Optional TimeOut As Long = PING_TIMEOUT) As Long

   Dim hPort As Long
   Dim dwAddress As Long
   Dim sDataToSend As String
   Dim iOpt As Long
   
   sDataToSend = "Echo This"
   dwAddress = AddressStringToLong(szAddress)
   
   hPort = IcmpCreateFile()
   
   If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), TimeOut) Then
   
        'the ping succeeded,
        '.Status will be 0
        '.RoundTripTime is the time in ms for
        '               the ping to complete,
        '.Data is the data returned (NULL terminated)
        '.Address is the Ip address that actually replied
        '.DataSize is the size of the string in .Data
         PingAddress = ECHO.RoundTripTime
   Else: PingAddress = ECHO.Status * -1
   End If
                       
   Call IcmpCloseHandle(hPort)
   
End Function
   

Private Function AddressStringToLong(ByVal tmp As String) As Long

Dim i As Integer
Dim parts(1 To 4) As String
   
    i = 0
    
    If InStr(1, tmp, ".", vbTextCompare) = 0 Then
        AddressStringToLong = gethostbyname(tmp)
    Else
        'we have to extract each part of the
        '123.456.789.123 string, delimited by
        'a period
        While InStr(tmp, ".") > 0
          i = i + 1
          parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
          tmp = Mid(tmp, InStr(tmp, ".") + 1)
        Wend
        
        i = i + 1
        parts(i) = tmp
        
        If i <> 4 Then
          AddressStringToLong = 0
          Exit Function
        End If
        
        'build the long value out of the
        'hex of the extracted strings
        AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & Right("00" & Hex(parts(3)), 2) & Right("00" & Hex(parts(2)), 2) & Right("00" & Hex(parts(1)), 2))
   End If
End Function


Private Sub SocketsCleanup()

   Dim X As Long
   
  'need to use a var (insread of embedding
  'in the If..Then call) becuse the function
  'returns the error code if failed.
   X = WSACleanup()

   If X <> 0 Then
       MsgBox "Windows Sockets error " & Trim$(Str$(X)) & " occurred in Cleanup.", vbExclamation
   End If
    
End Sub


Private Function SocketsInitialize() As Boolean

    Dim WSAD As WSAdata
    Dim X As Integer
    Dim szLoByte As String
    Dim szHiByte As String
    Dim szBuf As String
    
    X = WSAStartup(WS_VERSION_REQD, WSAD)
    
   'check for valid response
    If X <> 0 Then

        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
        Exit Function

    End If
    
   'check that the version of sockets is supported
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
       (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
        HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        
        szHiByte = Trim$(Str$(HiByte(WSAD.wVersion)))
        szLoByte = Trim$(Str$(LoByte(WSAD.wVersion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " is not supported by Windows " & _
                          "Sockets for 32 bit Windows environments."
        MsgBox szBuf, vbExclamation
        Exit Function
        
    End If
    
   'check that there are available sockets
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then

        szBuf = "This application requires a minimum of " & _
                 Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox szBuf, vbExclamation
        Exit Function

    End If
    
    SocketsInitialize = True
        
End Function

Public Function Ping(ByVal hostnameOrIpaddress As String, Optional timeOutmSec As Long = PING_TIMEOUT) As Boolean
Dim echoValues As ICMP_ECHO_REPLY
Dim pos As Integer
Dim Count As Integer
Dim returnIp As Collection
   
    On Error GoTo e_Trap
    If Trim(hostnameOrIpaddress) = "" Then
        Ping = False
        Exit Function
    End If
    
    If SocketsInitialize() Then
        
        If InStr(1, hostnameOrIpaddress, ".", vbTextCompare) <> 0 Then
            If IsNumeric(Mid(hostnameOrIpaddress, 1, InStr(1, hostnameOrIpaddress, ".") - 1)) = False Then
                Set returnIp = ResolveIpaddress(hostnameOrIpaddress)
                If returnIp.Count = 0 Then
                    Ping = False
                    Exit Function
                Else
                    hostnameOrIpaddress = returnIp.Item(1)
                End If
            End If
        End If
    
        'ping an ip address, passing the
        'address and the ECHO structure
        Call PingAddress((hostnameOrIpaddress), echoValues, timeOutmSec)
        
        If left$(echoValues.data, 1) <> chr$(0) Then
           pos = InStr(echoValues.data, chr$(0))
           echoValues.data = left$(echoValues.data, pos - 1)
        Else
              echoValues.data = ""
        End If
             
        SocketsCleanup
        
        If echoValues.Status <> 0 Then
            Ping = False
        Else
            Ping = True
        End If
    End If
    Exit Function
e_Trap:
    Ping = False
End Function

Public Function ResolveIpaddress(ByVal hostname As String) As Collection
Dim hostent_addr As Long
Dim Host As Hostent
Dim hostip_addr As Long
Dim temp_ip_address() As Byte
Dim i As Integer
Dim ip_address As String
Dim Count As Integer

    If SocketsInitialize() Then
    
        Set ResolveIpaddress = New Collection
        hostent_addr = gethostbyname(hostname)
        
        If hostent_addr = 0 Then
            SocketsCleanup
            Exit Function
        End If
        
        RtlMoveMemory Host, hostent_addr, LenB(Host)
        RtlMoveMemory hostip_addr, Host.hAddrList, 4
        
        'get all of the IP address if machine is  multi-homed
        
        Do
            ReDim temp_ip_address(1 To Host.hLength)
            RtlMoveMemory temp_ip_address(1), hostip_addr, Host.hLength
        
            For i = 1 To Host.hLength
                ip_address = ip_address & temp_ip_address(i) & "."
            Next
            ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
            ResolveIpaddress.Add ip_address
            ip_address = ""
            Host.hAddrList = Host.hAddrList + LenB(Host.hAddrList)
            RtlMoveMemory hostip_addr, Host.hAddrList, 4
         Loop While (hostip_addr <> 0)
    
        SocketsCleanup
    End If
End Function
Public Function ResolveHostname(ByVal ipaddress As String) As String

Dim hostip_addr As Long
Dim hostent_addr As Long
Dim newAddr As Long
Dim Host As Hostent
Dim strTemp As String
Dim strHost As String * 255

    If SocketsInitialize() Then
        newAddr = inet_addr(ipaddress)
        hostent_addr = gethostbyaddr(newAddr, Len(newAddr), AF_INET)

        If hostent_addr = 0 Then
            SocketsCleanup
            Exit Function
        End If

        RtlMoveMemory Host, hostent_addr, Len(Host)
        RtlMoveMemory ByVal strHost, Host.hName, 255
        strTemp = strHost
        If InStr(strTemp, chr(0)) <> 0 Then strTemp = left(strTemp, InStr(strTemp, chr(0)) - 1)
        strTemp = Trim(strTemp)
        ResolveHostname = strTemp
        SocketsCleanup

    End If
End Function
