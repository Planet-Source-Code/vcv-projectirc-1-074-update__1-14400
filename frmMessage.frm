VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Query 
   Caption         =   "Query"
   ClientHeight    =   3645
   ClientLeft      =   4935
   ClientTop       =   2070
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   6480
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
            TextRTF         =   $"frmMessage.frx":058A
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
         TextRTF         =   $"frmMessage.frx":0606
         MouseIcon       =   "frmMessage.frx":0682
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
            Picture         =   "frmMessage.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMessage.frx":0DF0
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   5
      Top             =   75
      Width           =   1425
   End
   Begin VB.Label lblHost 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ident@host"
      Height          =   195
      Left            =   1755
      TabIndex        =   4
      Top             =   75
      Width           =   825
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
Attribute VB_Name = "Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bControl As Boolean

Public strNick  As String
Public strHost  As String
Public newBuffer As Boolean
Dim bLink As Boolean
Dim strLink As String
Dim textHistory As New Collection
Dim intCurHist  As Integer
Public Sub PutText(strNick As String, strText As String)
    If left(strText, 8) = strAction & "ACTION " Then
        strText = RightR(strText, 8)
        strText = LeftR(strText, 1)
        PutData Me.DataIn, strColor & "06" & strNick & " " & strText
    ElseIf left(strText, 9) = strAction & "VERSION" & strAction Then
        Client.SendData "CTCPREPLY " & strNick & " VERSION :projectX for Windows9x"
    Else
        PutData Me.DataIn, Trim("" & strNick & " : " & chr(9) & strText)
    End If
End Sub


Sub DoHistory(strText As String)
    If textHistory.Count > MAX_TEXT_HISTORY Then
        textHistory.Remove 1
    End If
    textHistory.Add strText
    intCurHist = textHistory.Count + 1
End Sub


Sub DoConnect(strServer As String)
    Dim bDo As Integer
    
    bDo = MsgBox("Would you really like to connect to the server '" & strServer & "'? You will be disconnect from your current one", vbQuestion Or vbYesNo)
    
    If bDo = vbYes Then
        Client.mnu_File_Disconnect_Click
        Dim x As Integer
        GetAlias "", "server " & strServer
    End If
End Sub

Private Sub DataIn_Change()
    If Client.ActiveForm.Caption = Me.Caption Then
        newBuffer = False
    Else
        newBuffer = True
    End If
    Client.DrawToolbar
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
    ElseIf left(txt, 1) = "#" Then
        Dim inttemp As Integer
        inttemp = GetChanIndex(txt)
        If inttemp = -1 Then Client.SendData "JOIN " & txt
    ElseIf LCase(left(txt, 4)) = "irc." Then
        DoConnect txt
    Else
        Client.SendData "WHOIS " & strNick
        bLink = False
    End If

End Sub

Private Sub DataIn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim txt As String

    txt = RichWordOver(DataIn, x, y)
    If LCase(left(txt, 7)) = "http://" Or _
       LCase(left(txt, 6)) = "ftp://" Or _
       LCase(left(txt, 7)) = "mailto:" Or _
       LCase(left(txt, 4)) = "www." Or _
       LCase(Right(txt, 5)) = ".html" Or _
       LCase(Right(txt, 4)) = ".htm" _
    Then
        If DataIn.MousePointer <> 99 Then DataIn.MousePointer = 99
    ElseIf left(txt, 1) = "#" Then
        If DataIn.MousePointer <> 99 Then DataIn.MousePointer = 99
    ElseIf LCase(left(txt, 4)) = "irc." Then
        If DataIn.MousePointer <> 99 Then DataIn.MousePointer = 99
    Else
        If DataIn.MousePointer <> 1 Then DataIn.MousePointer = 1
    End If

End Sub


Private Sub DataIn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEvents
    Dim txt As String

    txt = RichWordOver(DataIn, x, y)
    strLink = txt
    
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
    
    Dim strData As String

    On Error Resume Next
    If KeyAscii = 13 Then
        strData = ANSICode(DataOut)
        If DataOut.Text = "" Then Exit Sub
        KeyAscii = 0
        
        If left(DataOut.Text, 1) = "/" Then
            Client.SendData GetAlias(strNick, RightR(strData, 1))
        Else
            Client.SendData "PRIVMSG " & Me.Caption & " :" & strData
            PutData DataIn, "" & strMyNick & " : " & chr(9) & strData
        End If
        DataOut.SelColor = lngForeColor
        DoHistory strData
        DataOut.Text = ""
        DataOut.Text = ""
        DataOut.SelColor = lngForeColor
        DataOut.SelBold = False
        DataOut.SelUnderline = False
        DataOut.SelStrikeThru = False
    End If
        
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


Private Sub DataOut_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu Client.mnu_Edit
End Sub


Private Sub Form_Activate()
    DataOut.SetFocus
    newBuffer = False
    Dim i As Integer
    i = GetWindowIndex(strNick)
    SetWinFocus i
    Client.intHover = -1
    Client.intActive = i
    Client.DrawToolbar
End Sub

Private Sub Form_Load()
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

        
    strPos = GetINI(winINI, "pos", "*" & strTemp, "-1,-1,-1,-1")
    
    If strPos = "-1,-1,-1,-1" Then
        Exit Sub
    End If
    
    strLst = Split(strPos, ",")
    
    On Error Resume Next
    Me.Move CInt(strLst(0)), CInt(strLst(1)), CInt(strLst(2)), CInt(strLst(3))


End Sub


Private Sub Form_Resize()
    
    'MsgBox Me.WindowState & "~" & vbMinimized
    If Me.WindowState = vbMinimized Then
        HideWin Me.Tag
        Exit Sub
    End If
    
    If Me.Width < 4500 Then Me.Width = 4500
    If Me.Height < 2500 Then Me.Height = 2500
    
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


Private Sub Form_Unload(Cancel As Integer)
    strNick = ""
    Me.Caption = ""
    lblNick = ""
    On Error Resume Next
    Unload Queries(Me.Tag)
    Unload Me
End Sub


