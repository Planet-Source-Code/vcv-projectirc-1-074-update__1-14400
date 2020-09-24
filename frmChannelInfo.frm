VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ChannelInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Channel Information for "
   ClientHeight    =   3795
   ClientLeft      =   2055
   ClientTop       =   1455
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   4650
      TabIndex        =   19
      Top             =   3060
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4650
      TabIndex        =   18
      Top             =   3420
      Width           =   795
   End
   Begin VB.CheckBox chkModeS 
      Caption         =   " &Secret"
      Height          =   255
      Left            =   2115
      TabIndex        =   16
      Top             =   3495
      Width           =   1980
   End
   Begin VB.CheckBox chkModeT 
      Caption         =   " &Topic set by OPs only "
      Height          =   255
      Left            =   75
      TabIndex        =   15
      Top             =   3495
      Width           =   1980
   End
   Begin VB.TextBox txtLimit 
      Height          =   315
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   14
      Top             =   2820
      Width           =   540
   End
   Begin VB.TextBox txtKey 
      Height          =   315
      Left            =   3000
      TabIndex        =   13
      Top             =   2490
      Width           =   1620
   End
   Begin VB.CheckBox chkModeL 
      Caption         =   " &Limit to"
      Height          =   255
      Left            =   2115
      TabIndex        =   12
      Top             =   2835
      Width           =   1395
   End
   Begin VB.CheckBox chkModeK 
      Caption         =   " &Key :"
      Height          =   255
      Left            =   2115
      TabIndex        =   11
      Top             =   2505
      Width           =   1980
   End
   Begin VB.CheckBox chkModeP 
      Caption         =   " &Private"
      Height          =   255
      Left            =   2115
      TabIndex        =   10
      Top             =   3165
      Width           =   1980
   End
   Begin VB.CheckBox chkModeM 
      Caption         =   " &Moderated"
      Height          =   255
      Left            =   75
      TabIndex        =   9
      Top             =   2835
      Width           =   1980
   End
   Begin VB.CheckBox chkModeI 
      Caption         =   " &Invite only"
      Height          =   255
      Left            =   75
      TabIndex        =   8
      Top             =   2505
      Width           =   1980
   End
   Begin VB.CheckBox chkModeN 
      Caption         =   " &No external messages "
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   3165
      Width           =   1980
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   345
      Left            =   4665
      TabIndex        =   6
      Top             =   2070
      Width           =   795
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   345
      Left            =   4665
      TabIndex        =   5
      Top             =   1710
      Width           =   795
   End
   Begin VB.CommandButton cmdInvites 
      Caption         =   "&Invites"
      Height          =   345
      Left            =   4665
      TabIndex        =   4
      Top             =   1155
      Width           =   795
   End
   Begin VB.CommandButton cmdExcepts 
      Caption         =   "E&xcepts"
      Height          =   345
      Left            =   4665
      TabIndex        =   3
      Top             =   795
      Width           =   795
   End
   Begin VB.CommandButton cmdBans 
      Caption         =   "&Bans"
      Height          =   345
      Left            =   4665
      TabIndex        =   2
      Top             =   435
      Width           =   795
   End
   Begin MSComctlLib.ListView lvBans 
      Height          =   2025
      Left            =   15
      TabIndex        =   1
      Top             =   420
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   3572
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bans"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "By"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "When"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfTopic 
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"frmChannelInfo.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "users"
      Height          =   195
      Left            =   3615
      TabIndex        =   17
      Top             =   2850
      Width           =   390
   End
End
Attribute VB_Name = "ChannelInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strChannel As String
Public strKey As String
Public lngLimt As Long

Public strModes As String
Public modeI As Boolean
Public modeK As Boolean
Public modeL As Boolean
Public modeM As Boolean
Public modeN As Boolean
Public modeP As Boolean
Public modeS As Boolean
Public modeT As Boolean

Public strMT As String
Sub AddBEI(strHost As String, strWho As String, strTime As String)
    If strHost = "" Then Exit Sub
    Dim lvCount As Integer, var As ListItem
    lvCount = lvBans.ListItems.Count + 1
        
    Set var = lvBans.ListItems.Add(, "", strHost, 0, 0)
    lvBans.ListItems(lvBans.ListItems.Count).ToolTipText = strTopic
    
    DoEvents
    On Error Resume Next
    var.SubItems(1) = strWho
    var.SubItems(2) = AscTime(CLng(strTime))

End Sub

Function ChangeMode(strMode As String, chk As CheckBox) As String
    Dim strPre As String
    
    'MsgBox chk.Value & "~" & strMode
    If chk.Value Then
        If InStr(strModes, strMode) Then
        Else
            ChangeMode = "+" & strMode
        End If
    Else
        If InStr(strModes, strMode) Then
            ChangeMode = "-" & strMode
        Else
        End If
    End If

End Function

Sub ParseTheModes()

    If InStr(strModes, "i") Then
        chkModeI.Value = 1
        modeI = True
    End If
    If InStr(strModes, "k") Then
        chkModeK.Value = 1
        modeK = True
    End If
    If InStr(strModes, "l") Then
        chkModeL.Value = 1
        modeL = True
    End If
    If InStr(strModes, "m") Then
        chkModeM.Value = 1
        modeM = True
    End If
    If InStr(strModes, "n") Then
        chkModeN.Value = 1
        modeN = True
    End If
    If InStr(strModes, "p") Then
        chkModeP.Value = 1
        modeP = True
    End If
    If InStr(strModes, "s") Then
        chkModeS.Value = 1
        modeS = True
    End If
    If InStr(strModes, "t") Then
        chkModeT.Value = 1
        modeT = True
    End If
    
End Sub


Private Sub cmdBans_Click()
    bGettingChanInfo = True
    lvBans.ListItems.Clear
    lvBans.ColumnHeaders.Item(1).Text = "Bans"
    Client.SendData "MODE " & strChannel & " b"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExcepts_Click()
    lvBans.ListItems.Clear
    lvBans.ColumnHeaders.Item(1).Text = "Excepts"
    bGettingChanInfo = True
End Sub


Private Sub cmdInvites_Click()
    lvBans.ListItems.Clear
    lvBans.ColumnHeaders.Item(1).Text = "Invites"
End Sub


Private Sub cmdOK_Click()
    Dim strNewMode As String, strExtra As String
    
    strMT = ""
    strNewMode = ""
    
    strNewMode = strNewMode & ChangeMode("i", chkModeI)
    strNewMode = strNewMode & ChangeMode("m", chkModeM)
    strNewMode = strNewMode & ChangeMode("n", chkModeN)
    strNewMode = strNewMode & ChangeMode("t", chkModeT)
    strNewMode = strNewMode & ChangeMode("k", chkModeK)
    If chkModeK Then
        If InStr(strModes, "k") Then Else strExtra = strExtra & " " & txtKey.Text
    Else
        If InStr(strModes, "k") Then strExtra = strExtra & " " & txtKey.Text
    End If
    
    strNewMode = strNewMode & ChangeMode("l", chkModeL)
    If chkModeL Then
        If InStr(strModes, "l") Then Else strExtra = strExtra & " " & txtLimit.Text
    Else
        If InStr(strModes, "l") Then strExtra = strExtra & " " & txtLimit.Text
    End If
    
    strNewMode = strNewMode & ChangeMode("p", chkModeP)
    strNewMode = strNewMode & ChangeMode("s", chkModeS)
    
    If strNewMode <> "" Then Client.SendData "mode " & strChannel & " " & strNewMode & strExtra
    
    Unload Me
    
End Sub

Private Sub cmdRemove_Click()
    If lvBans.ColumnHeaders.Item(1).Text = "Bans" Then
        Client.SendData "mode " & strChannel & " -b " & lvBans.SelectedItem.Text
    End If
End Sub

Private Sub rtfTopic_Change()
    rtfTopic.ToolTipText = rtfTopic.Text

End Sub

Private Sub rtfTopic_KeyDown(KeyCode As Integer, Shift As Integer)
    If rtfTopic.Tag = "locked" Then
    '    If KeyCode = 8 Then KeyCode = 0
    '    If KeyCode = 46 Then KeyCode = 0
    End If

End Sub


Private Sub rtfTopic_KeyPress(KeyAscii As Integer)
    'If rtbTopic.Tag = "locked" Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        Client.SendData "TOPIC " & strName & " :" & rtbTopic.Text
        KeyAscii = 0
    End If

End Sub

