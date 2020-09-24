VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form Status 
   Caption         =   "Status"
   ClientHeight    =   3630
   ClientLeft      =   7020
   ClientTop       =   1425
   ClientWidth     =   6540
   FillColor       =   &H00A27E66&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6540
   Begin HoverButton.Button btnConnect 
      Height          =   330
      Left            =   5715
      TabIndex        =   6
      Top             =   15
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   14268829
      Caption         =   "Button"
      CaptionDown     =   "Button"
      CaptionOver     =   "&Button"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Picture         =   "frmStatus.frx":058A
      Style           =   1
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   16
      IconWidth       =   16
   End
   Begin VB.PictureBox picFlat 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   75
      ScaleHeight     =   3195
      ScaleWidth      =   6405
      TabIndex        =   4
      Top             =   375
      Width           =   6405
      Begin VB.PictureBox picDO 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   6630
         TabIndex        =   5
         Top             =   2970
         Width           =   6630
         Begin RichTextLib.RichTextBox DataOut 
            Height          =   390
            Left            =   -45
            TabIndex        =   0
            Top             =   -45
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   688
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            TextRTF         =   $"frmStatus.frx":09DC
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
         Left            =   -15
         TabIndex        =   2
         Top             =   15
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   5159
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         Appearance      =   0
         TextRTF         =   $"frmStatus.frx":0A58
         MouseIcon       =   "frmStatus.frx":0AD4
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
   Begin HoverButton.Button btnDisconnect 
      Height          =   330
      Left            =   6120
      TabIndex        =   7
      Top             =   15
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "Button"
      CaptionDown     =   "Button"
      CaptionOver     =   "&Button"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Picture         =   "frmStatus.frx":0DEE
      Style           =   1
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   16
      IconWidth       =   16
   End
   Begin VB.Shape shpDI 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3225
      Left            =   1635
      Top             =   375
      Width           =   4875
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "not connected"
      Height          =   225
      Left            =   1740
      TabIndex        =   3
      Top             =   75
      Width           =   1095
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Server Status"
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
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   1425
   End
   Begin VB.Shape shpLeftC 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3930
      Left            =   -15
      Top             =   -15
      Width           =   1650
   End
End
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bControl As Boolean
Public newBuffer As Boolean
Private Sub Command1_Click()

End Sub


Sub Update()
    Dim mode As String
    mode = "[" & Client.ModeString & "] "
    If mode = "[] " Then mode = ""
    
    If Client.sock.State = 0 Then
        Status.Caption = "Status"
    Else
        Status.Caption = "Status : " & strMyNick & " " & mode & "on " & Client.sock.RemoteHost
    End If
End Sub


Private Sub btnConnect_Click()
    Client.mnu_File_Connect_Click
End Sub

Private Sub btnDisconnect_Click()
    Client.mnu_File_Disconnect_Click
End Sub


Private Sub DataIn_Change()
    'DataIn.SelStart = Len(DataIn.Text)
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

Private Sub DataIn_GotFocus()
    DoEvents
    If DataIn.SelLength = 0 Then DataOut.SetFocus
End Sub

Private Sub DataIn_LostFocus()
    Dim lngRet As Long
    lngRet = ShowCaret(DataIn.hWnd)
End Sub


Private Sub DataIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    DoEvents
    Dim txt As String

    txt = RichWordOver(DataIn, X, Y)
    If LCase(left(txt, 7)) = "http://" Or _
       LCase(left(txt, 6)) = "ftp://" Or _
       LCase(left(txt, 7)) = "mailto:" _
    Then
        If Button = 2 Then ShellExecute 0, "open", txt, "", "", 0
    End If
    DoEvents
    TimeOut 0.2
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

Private Sub DataIn_SelChange()
    'If DataIn.SelLength = 0 Then DataOut.SetFocus
    'Me.Caption = DataIn.SelLength
End Sub

Private Sub DataOut_Click()
    DoEvents
End Sub

Private Sub DataOut_GotFocus()
    DataOut.SelColor = lngForeColor
End Sub

Private Sub DataOut_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = True   'control

End Sub

Private Sub DataOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 11 Then
        ColorPicker.Move Client.left + Me.left + DataOut.left + 300, Client.top + Me.top + picDO.top + DataOut.Height + 910
        ColorPicker.Show
        If trans.fIsWin2000() Then trans.fSetTranslucency ColorPicker.hWnd, 190
        StayOnTop ColorPicker, True
        DataOut.SetFocus
    ElseIf IsNumeric(chr(KeyAscii)) Then
    Else
        ColorPicker.Hide
    End If
    
    Dim strData As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
       ' If InStr(DataOut.Text, vbCrLf) Then
        
        
        If tooltip.Visible = True Then
            tooltip.Hide
            tooltip.Visible = False
        End If
        strData = ANSICode(DataOut)
        If strData = "" Then Exit Sub
        DataOut.Text = ""
        
        If left(strData, 1) = "/" Then
            Client.SendData GetAlias("", RightR(strData, 1))
        Else
            Client.SendData strData
        End If
        tooltip.Hide
        tooltip.Visible = False
        DataOut.SelColor = lngForeColor
        DataOut.SelBold = False
        DataOut.SelUnderline = False
        DataOut.SelStrikeThru = False
        DataOut.SelColor = lngForeColor
    
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


Private Sub DataOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Client.mnu_Edit
End Sub


Private Sub Form_Activate()
    Dim lngRet As Long
    lngRet = HideCaret(DataIn.hWnd)
    
    On Error Resume Next
    DataIn.BackColor = lngBackColor
    DataOut.BackColor = lngBackColor
    DataOut.SetFocus
    newBuffer = False
    DoEvents
    Client.mnu_View_Status.Checked = True
    SetWinFocus 1
    Client.intHover = -1
    Client.intActive = 1
    Client.DrawToolbar
    DataOut.SetFocus
End Sub

Private Sub Form_GotFocus()
    If Me.Visible Then DataOut.SetFocus
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    '* only for win2k
    'MsgBox (100 - intTranslucency) * 2.5 & "~" & intTranslucency
    'If intTranslucency <> 0 Then Call trans.fSetTranslucency(Status.hWnd, (100 - intTranslucency) * 2.5)
    
    strBold = chr(bold)
    strUnderline = chr(underline)
    strColor = chr(Color)
    strReverse = chr(REVERSE)
    strAction = chr(ACTION)
    
    PutText Status.DataIn, "Welcome to " & strBold & strColor & "12" & "projectIRC" & strColor & "!"
    PutText Status.DataIn, "projectIRC version " & strColor & "04" & "1" & strColor & " build " & strColor & "04" & App.Revision '& strColor
    
    DoEvents
    Me.Visible = True
    TimeOut 0.3
    
    'For i = 1 To 100
    'PutTheData vbaRich, "Welcome to " & strBold & strColor & "12,15" & "projectIRC" & strColor & "!"
    'TimeOut 0.5
    'Next
    
    TimeOut 0.01
    
    '* Set the colors straight!!
    Me.BackColor = lngRightColor
    shpLeftC.BackColor = lngLeftColor
    shpDI.BorderColor = lngLeftColor
    picFlat.BackColor = lngLeftColor
    DataOut.SelColor = lngForeColor
    lbl1.ForeColor = lngBackColor
    shpTBorder.BorderColor = lngLeftColor
    SetButton btnConnect
    SetButton btnDisconnect
    
    '* Set the font!
    On Error Resume Next
    Status.DataIn.Font.Name = strFontName
    Status.DataIn.SelFontName = strFontName
    
    '* If connect on load, do it
    If bConOnLoad Then
        Call Client.mnu_File_Connect_Click
    End If
    
    'If BuddyList.Visible = False Then
    '    bShowBL = True
    '    Load BuddyList
    '    TimeOut 0.01
    'End If
    Me.Visible = True
    DoEvents
    Client.mnu_view_ResetAWPos_Click
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Me.Hide
End Sub

Private Sub Form_Resize()
    If Status.WindowState = vbMinimized Then Exit Sub
    If Status.Width < 4500 Then Me.Width = 4500
    If Status.Height < 2500 Then Status.Height = 2500
    
    '* Resizing.. oh boy
    If IFTiface = IFT_FANCY Then    'fancy interface
        shpDI.Width = Status.ScaleWidth - 1650
        shpDI.Height = Status.ScaleHeight - 410
        picFlat.Width = Status.ScaleWidth - 110
        picFlat.Height = Status.ScaleHeight - 425
        DataIn.Width = Status.ScaleWidth - 110
        DataIn.Height = Status.ScaleHeight - 700
        DataOut.Width = Status.ScaleWidth - 10
        picDO.top = Status.ScaleHeight - 680
        picDO.Width = Status.ScaleWidth - 120
        shpLeftC.Height = Status.ScaleHeight + 25
        
        btnConnect.left = Status.ScaleWidth - 810
        btnDisconnect.left = Status.ScaleWidth - 410
    Else    'simple - ignore this.. for now
        shpDI.top = 10
        shpDI.Width = Status.ScaleWidth - 1630
        shpDI.Height = Status.ScaleHeight - 10
        picFlat.Width = Status.ScaleWidth - 0
        DataIn.Width = Status.ScaleWidth - 0
        picFlat.Height = Status.ScaleHeight + 10
        DataIn.Height = Status.ScaleHeight - 270
        picFlat.top = -10
        picFlat.left = 0
        DataOut.Width = Status.ScaleWidth + 130
        DataOut.top = 0
        picDO.top = Status.ScaleHeight - 270
        picDO.Height = 430
        picDO.Width = Status.ScaleWidth - 0
        shpLeftC.Height = Status.ScaleHeight + 25
    End If
End Sub


Private Sub sock_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub


Private Sub picFlat_GotFocus()
    DataOut.SetFocus
End Sub

Private Sub picFlat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DataOut.SetFocus
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            '* Connect
            Client.mnu_File_Connect_Click
        Case 2
            '* Disconnect
            Call Client.mnu_File_Disconnect_Click
    End Select
End Sub


Private Sub Toolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DataOut.SetFocus
End Sub


