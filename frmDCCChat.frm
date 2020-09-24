VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DCCChat 
   Caption         =   "DCC Chat - ?"
   ClientHeight    =   3645
   ClientLeft      =   2325
   ClientTop       =   3540
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   6480
   Begin MSWinsockLib.Winsock sock 
      Left            =   1980
      Top             =   3330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picFlat 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   75
      ScaleHeight     =   3195
      ScaleWidth      =   6330
      TabIndex        =   0
      Top             =   390
      Width           =   6330
      Begin VB.PictureBox picDO 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   6360
         TabIndex        =   1
         Top             =   2955
         Width           =   6360
         Begin RichTextLib.RichTextBox DataOut 
            Height          =   390
            Left            =   -45
            TabIndex        =   2
            Top             =   -45
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   688
            _Version        =   393217
            MultiLine       =   0   'False
            TextRTF         =   $"frmDCCChat.frx":0000
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
         TabIndex        =   3
         Top             =   0
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   5159
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmDCCChat.frx":007C
         MouseIcon       =   "frmDCCChat.frx":00F8
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
      Left            =   3645
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDCCChat.frx":0412
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDCCChat.frx":0866
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpDI 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3240
      Left            =   1650
      Top             =   375
      Width           =   4785
   End
   Begin VB.Label lblHost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ident@host"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1755
      TabIndex        =   5
      Top             =   75
      Width           =   825
   End
   Begin VB.Label lblNick 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "nick"
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
      TabIndex        =   4
      Top             =   75
      Width           =   1425
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
Attribute VB_Name = "DCCChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strNick      As String
Public strHost      As String
Public lngPort      As Long
Public newBuffer    As Boolean
Public Sub PutText(strNick As String, strText As String)
    If Trim(strText) = "" Then Exit Sub
    If left(strText, 8) = strAction & "ACTION " Then
        strText = RightR(strText, 8)
        strText = LeftR(strText, 1)
        PutData Me.DataIn, strColor & "06" & strNick & " " & strText
    ElseIf left(strText, 9) = strAction & "VERSION" & strAction Then
        Client.SendData "CTCPREPLY " & strNick & " VERSION :" & strVersionReply & strAction
    Else
        PutData Me.DataIn, Trim("" & strNick & " : " & chr(9) & strText)
    End If
End Sub


Private Sub DataIn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu Client.mnu_Edit
End Sub


Private Sub DataOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 11 Then
        ColorPicker.Move Client.left + Me.left, Client.top + Me.top + Me.Height - 100
        ColorPicker.Show
    ElseIf IsNumeric(chr(KeyAscii)) Then
    Else
        ColorPicker.Hide
    End If
    
    Dim strData As String
    
    If KeyAscii = 13 Then
        strData = ANSICode(DataOut)
        If DataOut.Text = "" Then Exit Sub
        KeyAscii = 0
        
        If left(strData, 1) = "/" Then
            Client.SendData GetAlias(strNick, RightR(strData, 1))
        Else
            If sock.State = sckConnected Then sock.SendData strData & vbLf
            'PutData DataIn, "" & strMyNick & " : " & Chr(9) & DataOut.Text
            PutText strMyNick, strData
        End If
        DataOut.SelColor = lngForeColor
        'DoHistory DataOut.Text
        DataOut.Text = ""
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


Private Sub DataOut_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu Client.mnu_Edit
End Sub


Private Sub Form_Activate()
    DoEvents
    DataOut.SetFocus
    newBuffer = False
    Dim i As Integer
    i = GetWindowIndex(Me.Caption)
    SetWinFocus i
    Client.intHover = -1
    Client.intActive = i
    Client.DrawToolbar

End Sub


Private Sub Form_Load()
    DoEvents
'    Call MDIForm_Resize
    Client.DrawToolbar
    
    DataIn.BackColor = lngBackColor
    DataOut.BackColor = lngBackColor
    DataOut.SelColor = lngForeColor
    
    '* Set the colors straight!!
    Me.BackColor = lngRightColor
    shpLeftC.BackColor = lngLeftColor
    shpDI.BorderColor = lngLeftColor
    picFlat.BackColor = lngLeftColor
    
    Me.Visible = True
    DoEvents
    Client.mnu_view_ResetAWPos_Click
    
    Dim strTemp As String, strPos As String, strCPos As String, strLst() As String
    strTemp = strNick
    
    With Me
        strCPos = .left & "," & _
                 .top & "," & _
                 .Width & "," & _
                 .Height
    End With

        
    strPos = GetINI(winINI, "pos", "@" & strTemp, "-1,-1,-1,-1")
    
    If strPos = "-1,-1,-1,-1" Then
        Exit Sub
    End If
    
    strLst = Split(strPos, ",")
    
    On Error Resume Next
    Me.Move CInt(strLst(0)), CInt(strLst(1)), CInt(strLst(2)), CInt(strLst(3))

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    DoEvents
    If Val(Me.Tag) = intDCCChats Then intDCCChats = intDCCChats - 1
    Me.Tag = ""
    Cancel = 0
End Sub


Private Sub Form_Resize()
    
    shpDI.Width = Me.ScaleWidth - 1670
    shpDI.Height = Me.ScaleHeight - 410
    picFlat.Width = Me.ScaleWidth - 110
    picFlat.Height = Me.ScaleHeight - 450
    DataIn.Width = Me.ScaleWidth - 120
    DataIn.Height = Me.ScaleHeight - 710
    DataOut.Width = Me.ScaleWidth - 0
    picDO.top = Me.ScaleHeight - 690
    picDO.Width = Me.ScaleWidth - 120
    shpLeftC.Height = Me.ScaleHeight + 25


End Sub


Private Sub sock_Close()
    PutData DataIn, strColor & "02Connection was CLOSED"
End Sub

Private Sub sock_Connect()
    PutData DataIn, strColor & "02Connection was made, please feel free to now chat."
End Sub


Private Sub sock_ConnectionRequest(ByVal requestID As Long)
    sock.Close
    sock.Accept requestID
    
    PutData DataIn, strColor & "02* Accepted DCC Chat Connection from " & sock.RemoteHostIP & "."
    strHost = sock.RemoteHostIP
    lblHost = strHost
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, strLines() As String, i As Integer
    sock.GetData strData, vbString
    
    strLines = Split(strData, vbLf)
    For i = LBound(strLines) To UBound(strLines)
        PutText strNick, strLines(i)
        DoEvents
    Next i
End Sub


Private Sub sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData DataIn, strColor & "04ERROR occured in socket : " & Description
End Sub


