VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Channel 
   Caption         =   "#channel"
   ClientHeight    =   3660
   ClientLeft      =   2745
   ClientTop       =   1935
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChannel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6795
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   3630
      Left            =   6465
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   6403
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Click to Op/DeOp the selected user."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Click to give/take away the selected user helper status."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Click to Voice/DeVoice the selected user."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "t"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "i"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "l"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "m"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "n"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "p"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s"
            ImageIndex      =   13
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTopic 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1755
      ScaleHeight     =   255
      ScaleWidth      =   4665
      TabIndex        =   8
      Top             =   105
      Width           =   4665
      Begin RichTextLib.RichTextBox rtbTopic 
         Height          =   375
         Left            =   -45
         TabIndex        =   9
         Top             =   -45
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   661
         _Version        =   393217
         MultiLine       =   0   'False
         MaxLength       =   512
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmChannel.frx":058A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBMPC"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picFlat 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   75
      ScaleHeight     =   3195
      ScaleWidth      =   6330
      TabIndex        =   1
      Top             =   390
      Width           =   6330
      Begin VB.PictureBox picNicks 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   4680
         ScaleHeight     =   2925
         ScaleWidth      =   1650
         TabIndex        =   6
         Top             =   0
         Width           =   1650
         Begin VB.ListBox lstNicks 
            Height          =   2985
            IntegralHeight  =   0   'False
            Left            =   -30
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   -30
            Width           =   1710
         End
      End
      Begin VB.PictureBox picDO 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   6360
         TabIndex        =   2
         Top             =   2955
         Width           =   6360
         Begin RichTextLib.RichTextBox DataOut 
            Height          =   390
            Left            =   -45
            TabIndex        =   3
            Top             =   -45
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   688
            _Version        =   393217
            MultiLine       =   0   'False
            MaxLength       =   512
            TextRTF         =   $"frmChannel.frx":0606
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBMPC"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin RichTextLib.RichTextBox DataIn 
         Height          =   2925
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4655
         _ExtentX        =   8202
         _ExtentY        =   5159
         _Version        =   393217
         BorderStyle     =   0
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmChannel.frx":0682
         MouseIcon       =   "frmChannel.frx":06FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBMPC"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":0A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":0E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":12C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":15DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":18F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":1C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":1F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":224C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":2568
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":2884
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":2BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":2EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":31D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   3570
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpTopic 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   300
      Left            =   1740
      Top             =   90
      Width           =   4710
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "#channel name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   90
      TabIndex        =   5
      Top             =   75
      Width           =   1425
   End
   Begin VB.Shape shpDI 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3240
      Left            =   1650
      Top             =   375
      Width           =   4800
   End
   Begin VB.Shape shpLeftC 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3930
      Left            =   0
      Top             =   0
      Width           =   1650
   End
End
Attribute VB_Name = "Channel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strTopic As String
Public strMode  As String
Public strName  As String
Public strKey   As String
Public intLimit As Integer
Dim Nicks()     As Nick
Public intNicks As Integer
Public bControl As Boolean
Public bdown As Boolean
Dim modeS()     As typMode
Public intModes As Integer
Public newBuffer As Boolean
Dim bLink As Boolean
Dim strLink As String
Public intUserCount As Integer

Dim textHistory As New Collection
Dim intCurHist  As Integer
Public Sub AddMode(strMode As String, bPlus As Boolean)
    Dim i As Integer
    
    For i = 1 To Toolbar1.Buttons.Count
        If Toolbar1.Buttons.Item(i).Key = strMode Then
            Toolbar1.Buttons.Item(i).Value = tbrPressed
        End If
    Next i
    
    For i = 1 To intModes
        If modeS(i).mode = strMode Then Exit Sub
    Next i
    
    intModes = intModes + 1
    ReDim Preserve modeS(1 To intModes) As typMode
    
    With modeS(intModes)
        .bPos = True
        .mode = strMode
    End With
    Update
End Sub


Public Sub AddNick(strNick As String, Optional strHost As String = "")
    Dim strPre As String

    If strNick = "" Then Exit Sub
    intNicks = intNicks + 1
    ReDim Preserve Nicks(1 To intNicks) As Nick
    If InStr(strNick, "%") Then Nicks(intNicks).Helper = True: strPre = "%": strNick = Replace(strNick, "%", "")
    If InStr(strNick, "+") Then Nicks(intNicks).Voice = True: strPre = "+": strNick = Replace(strNick, "+", "")
    If InStr(strNick, "@") Then Nicks(intNicks).Op = True: strPre = "@": strNick = Replace(strNick, "@", "")
    
    Nicks(intNicks).Host = strHost
    Nicks(intNicks).Nick = strNick
    lstNicks.AddItem DisplayNick(Nicks(intNicks))
    intUserCount = intUserCount + 1
    Update
End Sub



Public Sub ChangeNck(strOldNick, strNewNick)
    Dim i As Integer, bInd As Integer
    
    For i = 1 To intNicks
        If LCase(Nicks(i).Nick) = LCase(strOldNick) Then
            Nicks(i).Nick = strNewNick
            bInd = i
            Exit For
        End If
    Next i
    For i = 0 To lstNicks.ListCount - 1
        If LCase(RealNick(lstNicks.List(i))) = LCase(strOldNick) Then
            lstNicks.List(i) = DisplayNick(Nicks(bInd))
            
            If LCase(strOldNick) = LCase(strNewNick) Then   'same nick, diff case
            'x has reformatted the capitalization in his or her nick to X

                PutData DataIn, strColor & "03" & strBold & strOldNick & strBold & " has reformatted the capitalization in his/her nick to " & strBold & strNewNick
            Else    'diff nick
                PutData DataIn, strColor & "03" & strBold & strOldNick & strBold & " is now known as " & strBold & strNewNick
            End If
            
            Exit For
        End If
    Next i
End Sub




Sub DoConnect(strServer As String)
    Dim bDo As Integer
    
    bDo = MsgBox("Would you really like to connect to the server '" & strServer & "'? You will be disconnect from your current one", vbQuestion Or vbYesNo)
    
    If bDo = vbYes Then
        Client.mnu_File_Disconnect_Click
        Dim X As Integer
        GetAlias "", "server " & strServer
    End If
End Sub

Sub DoHistory(strText As String)
    If textHistory.Count > MAX_TEXT_HISTORY Then
        textHistory.Remove 1
    End If
    textHistory.Add strText
    intCurHist = textHistory.Count + 1
End Sub

Public Function GetNick(intIndex As Integer) As String
    GetNick = Nicks(intIndex).Nick
End Function


Function InChannel(strNick As String) As Boolean
    Dim i As Integer
    For i = 1 To intNicks
        If LCase(strNick) = LCase(Nicks(i).Nick) Then InChannel = True: Exit Function
    Next i
    InChannel = False
End Function

Public Function isHalfOp(strNick As String) As Boolean
    Dim i As Integer
    For i = 1 To intNicks
        If LCase(Nicks(i).Nick) = LCase(strNick) Then
            If Nicks(i).Helper = True Then
                isHalfOp = True
                Exit Function
            Else
                isHalfOp = False
                Exit Function
            End If
        End If
    Next i
    isHalfOp = False
End Function

Public Function IsOp(strNick As String) As Boolean
    Dim i As Integer
    For i = 1 To intNicks
        If LCase(Nicks(i).Nick) = LCase(strNick) Then
            If Nicks(i).Op = True Then
                IsOp = True
                Exit Function
            Else
                IsOp = False
                Exit Function
            End If
        End If
    Next i
    IsOp = False
End Function

Public Function IsVoice(strNick As String) As Boolean
    Dim i As Integer
    For i = 1 To intNicks
        If LCase(Nicks(i).Nick) = LCase(strNick) Then
            If Nicks(i).Voice = True Then
                IsVoice = False
                Exit Function
            Else
                IsVoice = True
                Exit Function
            End If
        End If
    Next i
    IsVoice = False
End Function

Public Function ModeString() As String
    If intModes = 0 Then Exit Function
    Dim strFinal As String, bWhich As Boolean, i As Integer
    If modeS(1).bPos = True Then bWhich = True
    
    If bWhich Then strFinal = strFinal & "+" Else strFinal = strFinal & "-"
    
    For i = 1 To intModes
        If modeS(i).bPos <> bWhich Then
            bWhich = Not bWhich
            If bWhich Then strFinal = strFinal & "+" Else strFinal = strFinal & "-"
        End If
        strFinal = strFinal & modeS(i).mode
    Next i
    ModeString = strFinal
End Function

Function NickIndex(strNick As String) As Integer
    For i = 1 To intNicks
        If LCase(strNick) = LCase(Nicks(i).Nick) Then
            NickIndex = i
            Exit Function
        End If
    Next i
End Function

Public Sub PutText(strNick As String, strText As String)
    If left(strText, 1) = strAction Then
        HandleCTCP strNick, strText
    Else
        PutData Me.DataIn, Trim("" & strNick & " : " & chr(9) & strText)
    End If
End Sub


Public Sub HandleCTCP(strNick As String, strData As String)
    strData = RightR(strData, 1)
    strData = LeftR(strData, 1)
    
    Dim strCom As String, strParam As String, inttemp As Integer, strTemp As String, strArgs() As String
    Dim dccinfo As DCC_INFO
    
    Seperate strData, " ", strCom, strParam
    
    Select Case LCase(strCom)
        Case "version"
            PutData DataIn, strColor & "04" & "[" & strName & "] " & strCom
            PutData DataIn, strColor & "05" & strBold & strNick & strBold & " has just requested your client version"
            Client.CTCPReply strNick, "VERSION projectIRC for Windows"
        Case "ping"
            strTemp = RightOf(strData, " ")
            PutData DataIn, strColor & "04" & "[" & strName & "] " & strCom
            PutData DataIn, strColor & "05" & strBold & strNick & strBold & " has just pinged you"
            Client.CTCPReply strNick, "PING " & strTemp
        Case "action"
            PutData DataIn, strColor & "06" & strNick & " " & strParam
        Case "dcc"
            Seperate strData, " ", strTemp, strData
            strArgs = Split(strData, " ")
            
            Select Case LCase(strArgs(0))
                Case "send"
                    dccinfo.File = strArgs(1)
                    dccinfo.IP = strArgs(2)
                    
                    dccinfo.Port = strArgs(3)
                    dccinfo.Size = CLng(strArgs(4))
                    dccinfo.Nick = strNick
                    dccinfo.type = 2
                    inttemp = NewDCCSend(dccinfo)
                    TimeOut 0.5
                    DCCSends(inttemp).sock.Connect
            End Select
        Case Else
            PutData DataIn, strColor & "04" & "[" & strName & "] " & strCom
    End Select
End Sub



Public Sub RemoveMode(strMode As String)
    Dim i As Integer, j As Integer
    
    For i = 1 To Toolbar1.Buttons.Count
        If Toolbar1.Buttons.Item(i).Key = strMode Then
            Toolbar1.Buttons.Item(i).Value = tbrUnpressed
        End If
    Next i
    
    For i = 1 To intModes
        If modeS(i).mode = strMode Then
            modeS(i).mode = ""
            For j = i To intModes - 1
                modeS(j) = modeS(j + 1)
            Next j
            intModes = intModes - 1
            On Error Resume Next
            ReDim Preserve modeS(1 To intModes) As typMode
            Update
            Exit Sub
        End If
    Next i
End Sub

Public Sub RemoveNick(strNick As String)
    Dim i As Integer, j As Integer, strTemp As String
    
    For i = 1 To intNicks
        If Nicks(i).Nick = strNick Then
            For j = i To intNicks - 1
                Nicks(j) = Nicks(j + 1)
            Next j
            intNicks = intNicks - 1
            On Error Resume Next
            ReDim Preserve Nicks(1 To intNicks) As Nick
            
            For j = 0 To lstNicks.ListCount - 1
                strTemp = lstNicks.List(j)
                strTemp = Replace(strTemp, "@", "")
                strTemp = Replace(strTemp, "+", "")
                strTemp = Replace(strTemp, "%", "")
                If strTemp = strNick Then
                    lstNicks.RemoveItem j
                    intUserCount = intUserCount - 1
                    Update
                    Exit Sub
                End If
            Next j
            Exit Sub
        End If
    Next i
End Sub

Public Sub SetHelper(strNick As String, bWhich As Boolean)
    Dim i As Integer, bInd As Integer
    
    For i = 1 To intNicks
        If Nicks(i).Nick = strNick Then
            Nicks(i).Helper = bWhich
            bInd = i
            Exit For
        End If
    Next i
    For i = 0 To lstNicks.ListCount - 1
        If RealNick(lstNicks.List(i)) = strNick Then
            lstNicks.RemoveItem i
            lstNicks.AddItem DisplayNick(Nicks(bInd))
            Exit For
        End If
    Next i
End Sub

Public Sub SetOp(strNick As String, bWhich As Boolean)
    Dim i As Integer, bInd As Integer
    
    For i = 1 To intNicks
        If Nicks(i).Nick = strNick Then
            Nicks(i).Op = bWhich
            bInd = i
            Exit For
        End If
    Next i
    For i = 0 To lstNicks.ListCount - 1
        If RealNick(lstNicks.List(i)) = strNick Then
            lstNicks.RemoveItem i
            lstNicks.AddItem DisplayNick(Nicks(bInd))

            Exit For
        End If
    Next i
End Sub

Public Sub SetVoice(strNick As String, bWhich As Boolean)
    Dim i As Integer, bInd As Integer
    
    For i = 1 To intNicks
        If Nicks(i).Nick = strNick Then
            Nicks(i).Voice = bWhich
            bInd = i
            Exit For
        End If
    Next i
    For i = 0 To lstNicks.ListCount - 1
        If RealNick(lstNicks.List(i)) = strNick Then
            lstNicks.RemoveItem i
            lstNicks.AddItem DisplayNick(Nicks(bInd))

            Exit For
        End If
    Next i
End Sub

Sub Update()
    strMode = ModeString()
    Dim strExtra As String, strMd As String
    If intLimit <> 0 Then strExtra = strExtra & " " & CStr(intLimit)
    If strKey <> "" Then strExtra = strExtra & " " & strKey
    strMd = " [" & strMode & strExtra & "]"
    If strMd = " []" Then strMd = ""
    Me.Caption = strName & " [" & intUserCount & "]" & strMd
End Sub

Private Sub DataIn_Change()
    On Error Resume Next
    
    If Client.ActiveForm.Caption = Me.Caption Then
        newBuffer = False
    Else
        newBuffer = True
    End If
    Client.DrawToolbar
End Sub

Private Sub DataIn_Click()
    If DataIn.SelLength = 0 Then DataOut.SetFocus
End Sub

Private Sub DataIn_DblClick()
    Dim txt As String
    txt = strLink
    
    If LCase(left(txt, 7)) = "http://" Or _
       LCase(left(txt, 6)) = "ftp://" Or _
       LCase(left(txt, 7)) = "mailto:" Or _
       LCase(left(txt, 4)) = "www." Or _
       LCase(Right(txt, 5)) = ".html" Or _
       LCase(Right(txt, 4)) = ".htm" _
    Then
        ShellExecute 0, "open", txt, "", "", 0
    ElseIf InChannel(txt) Then
        DoEvents
        NewQuery txt, ""
        strLink = txt
    ElseIf left(txt, 1) = "#" Then
        Dim inttemp As Integer
        inttemp = GetChanIndex(txt)
        If inttemp = -1 Then Client.SendData "JOIN " & txt
    ElseIf LCase(left(txt, 4)) = "irc." Then
        DoConnect txt
    Else
        bLink = False
       
        ChannelInfo.strChannel = strName
        bGettingChanInfo = True
        ChannelInfo.rtfTopic.Text = ""
        ChannelInfo.Caption = "Channel Information for " & strName
        PutData ChannelInfo.rtfTopic, strTopic
        Client.SendData "MODE " & strName & " b"
        
        If IsOp(strMyNick) Then
            With ChannelInfo
                .rtfTopic.Locked = False
            End With
        Else
            With ChannelInfo
                .rtfTopic.BackColor = &H8000000F
                .rtfTopic.Locked = True
                .txtKey.Enabled = False
                .txtLimit.Enabled = False
                .cmdRemove.Enabled = False
                .cmdEdit.Enabled = False
            End With
            Dim X As Control
            For Each X In ChannelInfo
                If TypeOf X Is CheckBox Then
                    X.Enabled = False
                End If
            Next
            
        End If
        With ChannelInfo
            .strKey = strKey
            .txtKey = strKey
            If intLimit <> 0 Then .txtLimit = intLimit
            .lngLimt = CLng(intLimit)
            'MsgBox ModeString()
            .strModes = ModeString()
        End With
        
        TimeOut 0.01
        
'        MsgBox ChannelInfo.strModes
        
        PutData DataIn, strColor & "03* Collecting channel information"
        
        ChannelInfo.ParseTheModes
        ChannelInfo.Show vbModal

    End If
End Sub

Private Sub DataIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim txt As String

    txt = RichWordOver(DataIn, X, Y)
    If LCase(left(txt, 7)) = "http://" Or _
       LCase(left(txt, 6)) = "ftp://" Or _
       LCase(left(txt, 7)) = "mailto:" Or _
       LCase(left(txt, 4)) = "www." Or _
       LCase(Right(txt, 5)) = ".html" Or _
       LCase(Right(txt, 4)) = ".htm" _
    Then
        If DataIn.MousePointer <> 99 Then DataIn.MousePointer = 99
    ElseIf InChannel(txt) Then
        If DataIn.MousePointer <> 99 Then DataIn.MousePointer = 99
    ElseIf left(txt, 1) = "#" Then
        If DataIn.MousePointer <> 99 Then DataIn.MousePointer = 99
    ElseIf LCase(left(txt, 4)) = "irc." Then
        If DataIn.MousePointer <> 99 Then DataIn.MousePointer = 99
    Else
        If DataIn.MousePointer <> 1 Then DataIn.MousePointer = 1
    End If

End Sub


Private Sub DataIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEvents
    Dim txt As String

    txt = RichWordOver(DataIn, X, Y)
    strLink = txt
    If DataIn.SelLength = 0 Then DataOut.SetFocus
    
    If Button = 2 Then PopupMenu Client.mnu_Edit
End Sub

Private Sub DataOut_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = True   'control
    
    If KeyCode = 38 Then    'UP KEY!
        If intCurHist <= 1 Then Beep: Exit Sub
        intCurHist = intCurHist - 1
        DataOut.Text = textHistory.Item(intCurHist)
        KeyCode = 0
    ElseIf KeyCode = 40 Then    'down key!
        If intCurHist >= textHistory.Count Or intCurHist = -1 Then Beep: Exit Sub
        intCurHist = intCurHist + 1
        DataOut.Text = textHistory.Item(intCurHist)
        KeyCode = 0
    End If
End Sub

Private Sub DataOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 11 Then
        ColorPicker.Move Client.left + Me.left, Client.top + Me.top + Me.Height - 100
        ColorPicker.Show
    ElseIf IsNumeric(chr(KeyAscii)) Then
    Else
        ColorPicker.Hide
    End If
    
    Dim strText As String, strNick As String, i As Integer, bFound As Boolean
    Dim strData As String, strTemp As String
    
    bFound = False
    On Error Resume Next
    If KeyAscii = 13 Then
        strData = ANSICode(DataOut)
        If strData = "" Then Exit Sub
        KeyAscii = 0
        strText = strData
        
        If left(DataOut.Text, 1) = "/" Then
            strText = RightR(strText, 1)
            Client.SendData GetAlias(strName, strText)
            If Me.Tag = "PARTNOW" Then
                Me.Tag = "NOPART"
                Unload Me
                Exit Sub
            End If
        On Error Resume Next
        Else
            If bNickComplete Then
                Dim intPos1 As Integer, intPos2 As Integer
                strTemp = strText
                intPos1 = InStr(strText, ":")
                intPos2 = InStr(strText, " ")
                If intPos1 = 0 Then GoTo nonick
                If True Then
                    Seperate strText, ":", strNick, strText
                    strNick = LCase(Trim(strNick))
                    If strNick = "" Then GoTo blah
                    For i = 0 To lstNicks.ListCount - 1
                        If InStr(LCase(RealNick(lstNicks.List(i))), strNick) Then
                            bFound = True
                            strNick = RealNick(lstNicks.List(i))
                            strText = strNick & ": " & strText
                            GoTo nonick
                        End If
                    Next i
                End If
blah:
                strText = strTemp 'strNick & ":" & strText
            End If
nonick:

            Client.SendData "PRIVMSG " & strName & " :" & strText
            PutData DataIn, "" & strMyNick & " : " & chr(9) & strText
        End If
        DataOut.SelColor = lngForeColor
        DoHistory strData
        strData = ""
        DataOut.Text = ""
        DataOut.SelColor = lngForeColor
        DataOut.SelBold = False
        DataOut.SelUnderline = False
        DataOut.SelStrikeThru = False
    End If
    
    If bControl Then
        If KeyAscii = 11 Then
            DataOut.SelText = strColor
        ElseIf KeyAscii = 2 Then
            'DataOut.SelText = strBold
            DataOut.SelBold = Not DataOut.SelBold
        ElseIf KeyAscii = 21 Then
            'DataOut.SelText = strUnderline
            DataOut.SelUnderline = Not DataOut.SelUnderline
        ElseIf KeyAscii = 18 Then
            'DataOut.SelText = strReverse
            DataOut.SelStrikeThru = Not DataOut.SelStrikeThru
        End If
    End If
    
    If left(DataOut.Text, 1) = "/" Then
        strData = RightR(DataOut.Text, 1)
        If KeyAscii <> 8 And KeyAscii <> Asc(" ") Then strData = strData & chr(KeyAscii)
        strData = LeftOf(strData, " ")
        
        Dim strCommand As String, strInfo As String
        strCommand = strData
        
        'If KeyAscii <> 8 And KeyAscii <> Asc(" ") Then strCommand = strCommand & Chr(KeyAscii)
        'MsgBox KeyAscii
        
        
        DoTooltip strCommand, strInfo
        If strCommand = "" Then
            If tooltip.Visible = True Then
                tooltip.Hide
                tooltip.Visible = False
            End If
            Exit Sub
        End If
        
        If tooltip.lblCommand = strCommand And tooltip.Visible = True Then Exit Sub
        
        tooltip.lblCommand = strCommand
        tooltip.lblInfo = strInfo
        tooltip.Move Client.left + Me.left + DataOut.left + 300, Client.top + Me.top + picDO.top + DataOut.Height + 600
        tooltip.Show
        tooltip.Visible = True
        StayOnTop tooltip, True
        DataOut.SetFocus
        If trans.fIsWin2000() Then trans.fSetTranslucency tooltip.hWnd, 215
        
    Else
        If tooltip.Visible = True Then
            tooltip.Hide
            tooltip.Visible = False
        End If
    End If
    
End Sub


Private Sub DataOut_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = False   'control
End Sub

Private Sub DataOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bdown = True
End Sub

Private Sub DataOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bdown = True
    If Button = 2 Then PopupMenu Client.mnu_Edit
End Sub


Private Sub Form_Activate()
    If Me.WindowState = vbMinimized Or Me.Visible = False Then Else DataOut.SetFocus
    intCurHist = -1
    newBuffer = False
    Dim i As Integer
    i = GetWindowIndex(strName)
    SetWinFocus i
    Client.intActive = i
    Client.intHover = -1
    DoEvents
    Client.DrawToolbar
End Sub

Private Sub Form_GotFocus()
    DataOut.SetFocus
End Sub

Private Sub Form_Load()
        
    lstNicks.FontName = strFontName
    lstNicks.FontSize = intFontSize
    Client.DrawToolbar
    DataIn.BackColor = lngBackColor
    DataOut.BackColor = lngBackColor
    lstNicks.BackColor = lngBackColor
    rtbTopic.BackColor = lngBackColor
    rtbTopic.SelColor = lngForeColor
    DataOut.SelColor = lngForeColor
    lstNicks.ForeColor = lngForeColor
    
    '* Set the colors straight!!
    Me.BackColor = lngRightColor
    shpLeftC.BackColor = lngLeftColor
    shpTopic.BorderColor = lngLeftColor
    picTopic.BackColor = lngBackColor
    shpDI.BorderColor = lngLeftColor
    rtbTopic.SelColor = lngForeColor
    picFlat.BackColor = lngLeftColor
    On Error Resume Next
    Me.Visible = True
    DoEvents
    
    Dim strTemp As String, strPos As String, strCPos As String, strLst() As String
    strTemp = strName
    
    With Me
        strCPos = .left & "," & _
                 .top & "," & _
                 .Width & "," & _
                 .Height
    End With

    strPos = GetINI(winINI, "pos", "*" & strTemp, "-1,-1,-1,-1")
    If strPos = "-1,-1,-1,-1" Then
        Exit Sub
    End If
    
    strLst = Split(strPos, ",")
    
    On Error Resume Next
    Me.Move CInt(strLst(0)), CInt(strLst(1)), CInt(strLst(2)), CInt(strLst(3))


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    If Me.Tag <> "NOPART" Then
        Client.SendData "PART " & strName & " :closed channel"
    End If
    Me.Tag = ""
    
    strName = ""
    Me.Caption = ""
    lblName = ""
    strMode = ""
    intModes = 0
    intUserCount = 0
    intCurHist = 0
    
    Dim i As Integer
    For i = 1 To textHistory.Count
        textHistory.Remove 1
    Next i
    
    strKey = ""
    intLimit = 0
    On Error Resume Next
    If Val(Me.Tag) = intChannels Then intChannels = intChannels - 1
    Unload Channels(Me.Tag)
    Unload Me
    Cancel = 0
    Client.DrawToolbar
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then
        HideWin GetWindowIndex(strName)
        Exit Sub
    End If
    
    If Me.Width < 4500 Then Me.Width = 4500
    If Me.Height < 2500 Then Me.Height = 2500
    
    shpTopic.Width = Me.ScaleWidth - 2130
    picTopic.Width = shpTopic.Width - 50
    rtbTopic.Width = picTopic.Width + 100
    shpDI.Width = Me.ScaleWidth - 2040
    shpDI.Height = Me.ScaleHeight - 410
    picFlat.Width = Me.ScaleWidth - 500
    picFlat.Height = Me.ScaleHeight - 420
    DataIn.Width = Me.ScaleWidth - 2160
    DataIn.Height = Me.ScaleHeight - 700
    DataOut.Width = Me.ScaleWidth - 410
    picNicks.left = Me.ScaleWidth - 2140
    picNicks.Height = DataIn.Height
    lstNicks.Height = DataIn.Height + 60
    picDO.top = Me.ScaleHeight - 690
    picDO.Width = Me.ScaleWidth - 150
    shpLeftC.Height = Me.ScaleHeight + 25
    Toolbar.left = Me.ScaleWidth - 820
    Toolbar1.left = Me.ScaleWidth - 350
End Sub


Private Sub lstNicks_DblClick()
    Dim strNck As String, strHst As String
    strNck = RealNick(lstNicks.List(lstNicks.ListIndex))
    strHst = Nicks(NickIndex(strNck)).Host
    NewQuery strNck, strHst
End Sub


Private Sub lstNicks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        
        Dim strNick As String
        strNick = RealNick(lstNicks.List(lstNicks.ListIndex))
        
        If IsOp(strNick) Then
            Client.mnu_nicks_op.Caption = "De&Op"
        Else
            Client.mnu_nicks_op.Caption = "&Op"
        End If
        If IsVoice(strNick) Then
            Client.mnu_nicks_voice.Caption = "&Voice"
        Else
            Client.mnu_nicks_voice.Caption = "De&Voice"
        End If
        If isHalfOp(strNick) Then
            Client.mnu_nicks_halfop.Caption = "De&HalfOp"
        Else
            Client.mnu_nicks_halfop.Caption = "&HalfOp"
        End If
        Dim intHeight As Integer
        Me.FontName = lstNicks.FontName
        DoEvents
        intHeight = Me.TextHeight("ABCabcWyZyXx")
        PopupMenu Client.mnu_nicks, , lstNicks.left + picNicks.left + 100, picNicks.top + lstNicks.top + 410 + ((lstNicks.ListIndex - lstNicks.TopIndex + 1) * (intHeight + 30)) + 15
    End If
End Sub


Private Sub rtbTopic_Change()
    rtbTopic.ToolTipText = rtbTopic.Text
End Sub

Private Sub rtbTopic_KeyDown(KeyCode As Integer, Shift As Integer)
    If rtbTopic.Tag = "locked" Then
        If KeyCode = 8 Then KeyCode = 0
        If KeyCode = 46 Then KeyCode = 0
    End If
End Sub

Private Sub rtbTopic_KeyPress(KeyAscii As Integer)
    If rtbTopic.Tag = "locked" Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        Client.SendData "TOPIC " & strName & " :" & rtbTopic.Text
        KeyAscii = 0
    End If
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strNick As String
    strNick = RealNick(lstNicks.List(lstNicks.ListIndex))
    
    If IsOp(strNick) Then
        Client.mnu_nicks_op.Caption = "De&Op"
    Else
        Client.mnu_nicks_op.Caption = "&Op"
    End If
    If IsVoice(strNick) Then
        Client.mnu_nicks_voice.Caption = "&Voice"
    Else
        Client.mnu_nicks_voice.Caption = "De&Voice"
    End If
    If isHalfOp(strNick) Then
        Client.mnu_nicks_halfop.Caption = "De&HalfOp"
    Else
        Client.mnu_nicks_halfop.Caption = "&HalfOp"
    End If
    

    Select Case Button.Index
        Case 1:              Client.mnu_nicks_op_Click
        Case 2:              Client.mnu_nicks_halfop_Click
        Case 3:              Client.mnu_nicks_voice_Click
        Case Else:
            Dim strPlus As String, strAppend As String
            strAppend = ""
            strPlus = "-"
            
            If Button.Value = tbrPressed Then strPlus = "+"
            
            
            If Button.Key = "k" And Button.Value = tbrPressed Then
                strAppend = InputBox("Please enter the new channel key:", "Change key", strKey)
                If strAppend = "" Then Exit Sub
                'Button.Value = tbrUnpressed
            ElseIf Button.Key = "l" And Button.Value = tbrPressed Then
                strAppend = InputBox("Enter the new channels user limit:", "Change user limit", intLimit)
                If strAppend = "" Then Exit Sub
                'Button.Value = tbrUnpressed
            ElseIf Button.Key = "k" And Button.Value = tbrUnpressed Then
                strAppend = strKey
                'Button.Value = tbrPressed
            End If
            If Button.Value = tbrPressed Then
                Button.Value = tbrUnpressed
            Else
                Button.Value = tbrPressed
            End If
            
            If strAppend <> "" Then strAppend = " " & strAppend
            Client.SendData "MODE " & strName & " " & strPlus & Button.Key & strAppend
    End Select
End Sub

