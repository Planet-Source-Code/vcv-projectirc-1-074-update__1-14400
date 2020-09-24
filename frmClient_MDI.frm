VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.MDIForm Client 
   BackColor       =   &H8000000C&
   Caption         =   "projectIRC"
   ClientHeight    =   4680
   ClientLeft      =   2505
   ClientTop       =   2580
   ClientWidth     =   7740
   Icon            =   "frmClient_MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Begin MSScriptControlCtl.ScriptControl cScript 
      Left            =   1140
      Top             =   2355
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1725
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":1496
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":1D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":21C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":261A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":3C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":44FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTask 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   516
      TabIndex        =   1
      Top             =   4350
      Width           =   7740
      Begin VB.PictureBox picTaskBuffer 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   525
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   732
         TabIndex        =   2
         Top             =   540
         Visible         =   0   'False
         Width           =   10980
      End
      Begin VB.Timer tmrTask 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1755
         Top             =   30
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1155
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":51D6
            Key             =   ""
            Object.Tag             =   "&Connect"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":798A
            Key             =   ""
            Object.Tag             =   "&Disconnect"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":A13E
            Key             =   ""
            Object.Tag             =   "&Options"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":C8F2
            Key             =   ""
            Object.Tag             =   "&Join Channel"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":CE8E
            Key             =   ""
            Object.Tag             =   "&Open Query"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":D42A
            Key             =   ""
            Object.Tag             =   "&StatusWin"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":D9C6
            Key             =   ""
            Object.Tag             =   "&Show BuddyList"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":1017A
            Key             =   ""
            Object.Tag             =   "&Download"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolMain 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   516
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      Begin VB.PictureBox picBackTemp 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5955
         ScaleHeight     =   300
         ScaleWidth      =   510
         TabIndex        =   8
         Top             =   45
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picBack 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   5250
         ScaleHeight     =   150
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   75
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picBGImage 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   9930
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   9930
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         ForeColor       =   &H00FF00FF&
         Height          =   300
         Left            =   6975
         ScaleHeight     =   480
         ScaleMode       =   0  'User
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   90
         Visible         =   0   'False
         Width           =   300
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   2190
         TabIndex        =   3
         Top             =   60
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Click here to connect"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Click to here to disconnect from the server"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Click here for options"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Click here to join a channel"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Click here to query someone"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   330
         Left            =   195
         TabIndex        =   5
         Top             =   60
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   7
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   8
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   10000
         Y1              =   1
         Y2              =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   10000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   1280
         Y1              =   28
         Y2              =   28
      End
   End
   Begin MSWinsockLib.Winsock IDENT 
      Left            =   1155
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   113
      LocalPort       =   113
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   1575
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_Connect 
         Caption         =   "&Connect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_File_Disconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_File_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Options 
         Caption         =   "&Options..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_ScriptAliases 
         Caption         =   "&Script Aliases..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_File_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_minSystray 
         Caption         =   "&Minimize to System Tray"
      End
      Begin VB.Menu mnu_File_Quit 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu_Edit_Undo 
         Caption         =   "U&ndo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_Edit_lb00 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnu_Edit_Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnu_Edit_Paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnu_Edit_Delete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnu_Edit_lb01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_selectall 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu_Edit_Lcase 
         Caption         =   "&LowerCase Selected"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu_Edit_UCase 
         Caption         =   "&UpperCase Selected"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnu_view 
      Caption         =   "&View"
      Begin VB.Menu mnu_View_BuddyList 
         Caption         =   "&Buddy List"
      End
      Begin VB.Menu mnu_View_Debug 
         Caption         =   "&Debug Window"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_View_Status 
         Caption         =   "&Status Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_view_lb01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_view_TBTop 
         Caption         =   "Taskbar on &Top"
      End
      Begin VB.Menu mnu_View_TBBot 
         Caption         =   "Taskbar on &Bottom"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_View_TBrTop 
         Caption         =   "Toolbar on Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_View_TBrBot 
         Caption         =   "Toolbar on Bottom"
      End
      Begin VB.Menu mnu_view_lb02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_view_3darea 
         Caption         =   "&3D Client Area"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_view_flatarea 
         Caption         =   "&Flat Client Area"
      End
   End
   Begin VB.Menu mnu_commands 
      Caption         =   "&Commands"
      Begin VB.Menu mnu_commands_join 
         Caption         =   "&Join..."
         Shortcut        =   ^J
      End
      Begin VB.Menu mnu_commands_Part 
         Caption         =   "&Part"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_commands_lb01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_commands_query 
         Caption         =   "&Query User..."
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnu_commands_noticeuser 
         Caption         =   "&Notice User..."
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnu_window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_Window_Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnu_Window_TileH 
         Caption         =   "&Tile Horizontally"
      End
      Begin VB.Menu mnu_Tile_Vertically 
         Caption         =   "&Tile Vertically"
      End
      Begin VB.Menu mnu_windows_lb01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_view_RemClientPos 
         Caption         =   "&Remember Client Position"
      End
      Begin VB.Menu mnu_view_ForClientPos 
         Caption         =   "&Forget Client Position"
      End
      Begin VB.Menu mnu_view_ResClientPos 
         Caption         =   "&Reset Client Position"
      End
      Begin VB.Menu mnu_window_lb02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_view_RemAWPos 
         Caption         =   "&Remember Active Window Position"
      End
      Begin VB.Menu mnu_view_ForAWPos 
         Caption         =   "&Forget Active Window Position"
      End
      Begin VB.Menu mnu_view_ResetAWPos 
         Caption         =   "&Reset Active Window Position"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_Help_About 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnu_nicks 
      Caption         =   "nicks"
      Visible         =   0   'False
      Begin VB.Menu mnu_nicks_WhoIs 
         Caption         =   "&WhoIs"
      End
      Begin VB.Menu mnu_nicks_LB05 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_nicks_op 
         Caption         =   "&Op"
      End
      Begin VB.Menu mnu_nicks_halfop 
         Caption         =   "&HalfOp"
      End
      Begin VB.Menu mnu_nicks_voice 
         Caption         =   "&Voice"
      End
      Begin VB.Menu mnu_nicks_lb01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_nicks_Kick 
         Caption         =   "&Kick"
      End
      Begin VB.Menu mnu_nicks_Ban 
         Caption         =   "&Ban"
      End
      Begin VB.Menu mnu_nicks_lb02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_nicks_dccchat 
         Caption         =   "DCC &Chat"
      End
      Begin VB.Menu mnu_nicks_dccsend 
         Caption         =   "DCC &Send"
      End
      Begin VB.Menu mnu_nicks_lb03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_nicks_CTCP 
         Caption         =   "CTCP"
         Begin VB.Menu mnu_nicks_CTCP_PING 
            Caption         =   "&PING"
         End
         Begin VB.Menu mnu_nicks_CTCP_TIME 
            Caption         =   "&TIME"
         End
         Begin VB.Menu mnu_nicks_CTCP_VERSION 
            Caption         =   "&VERSION"
         End
      End
   End
   Begin VB.Menu mnu_Systray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnu_Systray_Restore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnu_Systray_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Systray_Quit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const clrSep = &H80000010

Public intActive As Integer, sngLastX As Integer
Public intHover As Integer, bReDraw As Boolean, intLast As Integer

Const buttonShadow = &H80000010
Const buttonHilight = &H80000014

Dim modeS()     As typMode
Public intModes As Integer

Public strDataBuffer As String, minSize As Long
Dim strHoverCap As String
Public SysTray As New CSystrayIcon

Public objScript As New clsScript

Public intCTries As Integer
'Public strServer As String
Sub ClearModes()
    For i = 1 To intModes
        modeS(i).bPos = False
        modeS(i).mode = ""
    Next i
    intModes = 0
End Sub


Sub InitializeScript()
    cScript.AddObject "Client", objScript, True
    objScript.Active = "@"
End Sub

Sub LoadAliases()
    
    Dim strData As String, strName As String, strCode As String, strLst() As String, i As Integer
    
    On Error Resume Next
    If FileExists(path & "aliases.data") = False Then
        Open path & "aliases.data" For Output As #1
            Print #1, ""
        Close #1
    End If
    
    Open path & "aliases.data" For Binary As #1
        strData = String(LOF(1), 0)
        Get #1, , strData
    Close #1
    
    If Err Then
        MsgBox "An error occured while trying to load the aliases." & vbCrLf & _
               "ERROR #" & Err & " : " & Error, vbCritical
        Exit Sub
    End If
    
    strLst = Split(strData, chr(0))
    strData = ""
    
    For i = LBound(strLst) To UBound(strLst)
        Seperate strLst(i), chr(8), strName, strCode
        If Replace(strName, chr(0), "") = "" Then Exit Sub
        
        AddAlias strName, strCode
    Next i
    
    strName = ""
    strCode = ""
    ReDim strLst(1)
    
End Sub

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

Public Sub AddMode(strMode As String, bPlus As Boolean)
    Dim i As Integer
    For i = 1 To intModes
        If modeS(i).mode = strMode Then Exit Sub
    Next i
    
    intModes = intModes + 1
    ReDim Preserve modeS(1 To intModes) As typMode
    
    With modeS(intModes)
        .bPos = True
        .mode = strMode
    End With
    Status.Update
End Sub

Sub PutCTCPReply(strNick As String, strKind As String, strReply As String)
    Dim strFinal As String, lngTemp As Long
    
    If LCase(strKind) = "ping" Then
        If IsNumeric(strReply) Then
            lngTemp = lngPingReply - CLng(strReply)
            strReply = Duration(lngTemp / 1000) & " " & lngTemp Mod 1000 & " milliseconds"
        End If
    End If
    
    strFinal = strColor & "04CTCPReply " & strBold & "[" & strBold & strNick & strUnderline & "!" & strUnderline & strKind & strBold & "]" & strBold & " " & strReply
    
    On Error Resume Next
    PutData Client.ActiveForm.DataIn, strFinal
    If Err Then
        PutData Status.DataIn, strFinal
    End If
End Sub

Public Sub RemoveMode(strMode As String)
    Dim i As Integer, j As Integer
    
    For i = 1 To intModes
        If modeS(i).mode = strMode Then
            modeS(i).mode = ""
            For j = i To intModes - 1
                modeS(j) = modeS(j + 1)
            Next j
            intModes = intModes - 1
            If intModes > 0 Then ReDim Preserve modeS(1 To intModes) As typMode
            Status.Update
            Exit Sub
        End If
    Next i
    Status.Update
End Sub

Sub CTCPReply(strNick As String, strReply As String)
    Client.SendData "NOTICE " & strNick & " :" & strAction & strReply & strAction
End Sub

Sub DrawToolbar()
    Dim intSeps As Integer, CenX As Integer, j As Integer, realWidth As Long
    Dim strTitle As String, intWidth As Integer, i As Integer, intBegin As Integer
    Dim intEnd As Integer, strDrawText As String, lngRet As Long, iconY As Integer, iconX As Integer
    Dim intStartY As Integer, strPerc As String, intPerc As Integer, dblPerc As Double
    Dim drawbevel As Boolean
    intStartY = 1
    bReDraw = True
    
    intSeps = WindowCount - 1
    
    If bStretchButtons Then
        realWidth = Client.ScaleWidth / 15 - 1
        minSize = (intSeps + 1) * (ICON_SIZE + 40)
        If realWidth < minSize Then realWidth = minSize Else minSize = realWidth
    Else
        realWidth = (intSeps + 1) * intButtonWidth
        minSize = realWidth
    End If
    
    picTaskBuffer.Width = realWidth
    
    intWidth = realWidth / (intSeps + 1)
    picTaskBuffer.Cls
    
    picTaskBuffer.CurrentY = intStartY
    '* If one window open, draw only single thing..
    If intSeps = 0 Then
        strTitle = GetWindowTitle(1)
        strDrawText = TaskText(intWidth, strTitle)
        CenX = (realWidth - picTaskBuffer.TextWidth(strDrawText)) / 2
        picTaskBuffer.CurrentX = CenX
        picTaskBuffer.ForeColor = vbRed
        picTaskBuffer.Print strDrawText;
        Exit Sub
    End If
    
    picTaskBuffer.ForeColor = clrSep
    For i = 1 To intSeps + 5
        picTaskBuffer.CurrentX = intWidth * i
        For j = 3 To picTaskBuffer.ScaleHeight - 4 Step 2
            picTaskBuffer.PSet (picTaskBuffer.CurrentX, j)
        Next j
    Next i
    
    picTaskBuffer.ForeColor = vbBlack
    For i = 1 To intSeps + 1
        picTaskBuffer.CurrentY = intStartY
        picTaskBuffer.CurrentX = intWidth * (i - 1)
        intBegin = picTaskBuffer.CurrentX + 2
        intEnd = intBegin + intWidth - 2
        
        strTitle = GetWindowTitle(Int(i))
        On Error Resume Next
        picIcon.Picture = LoadPicture("")
        
        If i = intActive Then
            picIcon.BackColor = &H80000000
        Else
            picIcon.BackColor = &H8000000F
        End If
        
        If strTitle Like "Status" Then
            picIcon.Picture = ImageList1.ListImages.Item(6).Picture
        ElseIf strTitle Like "Friend Tracker" Then
            picIcon.Picture = ImageList1.ListImages.Item(7).Picture
        ElseIf strTitle Like "[#]*" Then
            picIcon.Picture = ImageList1.ListImages.Item(4).Picture
        ElseIf strTitle Like "DCC Send*" Then
            picIcon.Picture = ImageList1.ListImages.Item(8).Picture
        Else
            picIcon.Picture = ImageList1.ListImages.Item(5).Picture
        End If
        
        picIcon.Picture = picIcon.Image
        
        'DoEvents
        intBegin = intBegin + 1
        intEnd = intEnd - 1
        
        If intActive = i Then iconY = 4 Else iconY = 3
        If intActive = i Then iconX = intBegin + 3 Else iconX = intBegin + 2
        
        '* Draw flat?
        If bFlatButtons Then
            If intHover = i Then
                drawbevel = True
            Else
                drawbevel = False
            End If
        Else
            drawbevel = True
        End If
        If WindowNewBuffer(i) Then drawbevel = False
        
        intEnd = intEnd - 2
        If intHover = i Then strHoverCap = " " & strTitle & " "
        If intActive = i Then
            picTaskBuffer.ForeColor = vbRed
            '* Active window, let's draw a inset bevel
            picTaskBuffer.Line (intBegin, intStartY)-(intEnd, picTaskBuffer.ScaleWidth - 2), &H80000000, BF       '&H80000000
            'picTaskBuffer.Line (intBegin, intStartY)-(intEnd, intStartY), buttonShadow
            'picTaskBuffer.Line (intBegin, intStartY)-(intBegin, picTaskBuffer.ScaleHeight - 2), buttonShadow
            
            picTaskBuffer.Line (intEnd, intStartY)-(intEnd, picTaskBuffer.ScaleHeight - 1), buttonHilight
            picTaskBuffer.Line (intBegin, picTaskBuffer.ScaleHeight - 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), buttonHilight
            
            picTaskBuffer.Line (intBegin, intStartY)-(intEnd + 1, intStartY), vbBlack
            picTaskBuffer.Line (intBegin, intStartY)-(intBegin, picTaskBuffer.ScaleHeight - 1), vbBlack
            
            picTaskBuffer.Line (intBegin + 1, intStartY + 1)-(intEnd, intStartY + 1), buttonShadow
            picTaskBuffer.Line (intBegin + 1, intStartY + 1)-(intBegin + 1, picTaskBuffer.ScaleHeight - 2), buttonShadow
            
            picTaskBuffer.CurrentY = intStartY + 1
            If intHover = i Then picTask.Tag = intHover
        ElseIf drawbevel Then
            
            picTask.Tag = intHover
            picTaskBuffer.ForeColor = vbBlack
            picTaskBuffer.Line (intBegin, intStartY)-(intEnd, intStartY), buttonHilight
            picTaskBuffer.Line (intBegin, intStartY)-(intBegin, picTaskBuffer.ScaleHeight - 2), buttonHilight
            picTaskBuffer.Line (intEnd, intStartY)-(intEnd, picTaskBuffer.ScaleHeight - 1), buttonShadow
            picTaskBuffer.Line (intBegin, picTaskBuffer.ScaleHeight - 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), buttonShadow
            picTaskBuffer.CurrentY = intStartY
        ElseIf WindowNewBuffer(i) Then
            Dim clr As Long
            clr = vbRed
            picTaskBuffer.Line (intBegin, intStartY)-(intEnd, intStartY), clr
            picTaskBuffer.Line (intBegin, intStartY)-(intBegin, picTaskBuffer.ScaleHeight - 2), clr
            picTaskBuffer.Line (intEnd, intStartY)-(intEnd, picTaskBuffer.ScaleHeight - 2), clr
            picTaskBuffer.Line (intBegin, picTaskBuffer.ScaleHeight - 2)-(intEnd, picTaskBuffer.ScaleHeight - 2), clr
            picTaskBuffer.ForeColor = vbBlack
            picTaskBuffer.CurrentY = intStartY
        Else
            If picTaskBuffer.ForeColor <> vbBlack Then picTaskBuffer.ForeColor = vbBlack
            picTaskBuffer.CurrentY = intStartY
        End If
        
        If Right(strTitle, 1) = "%" Then    'dcc transfer - show percent
            intBegin = intBegin + 18
            strPerc = Replace(RightOf(strTitle, " - "), "-", "")
            strPerc = LeftR(strPerc, 1)
            dblPerc = CDbl(strPerc)
            dblPerc = intBegin + (((intWidth - 18) / 100) * dblPerc)
            If dblPerc < intBegin Then dblPerc = intBegin + 1
            If dblPerc >= intEnd - 2 Then dblPerc = intEnd - 2
            If CInt(strPerc) = 0 Then Else picTaskBuffer.Line (intBegin + 2, intStartY + 2)-(CInt(dblPerc), picTaskBuffer.ScaleHeight - 4), &H8000000D, BF
            picTaskBuffer.ForeColor = &HFF00FF
            picTaskBuffer.Refresh
            picTaskBuffer.CurrentY = intStartY
            If intActive = i Then picTaskBuffer.CurrentY = picTaskBuffer.CurrentY + 1
        End If
        
        '* Let's put the icon
        
        lngRet = BitBlt(picTaskBuffer.hDC, iconX, iconY, 16, 16, picIcon.hDC, 0, 0, SRCCOPY)
        
        picTaskBuffer.CurrentX = intWidth * (i - 1)
        intBegin = picTaskBuffer.CurrentX
        intEnd = intBegin + intWidth
        
        strDrawText = TaskText(intWidth, strTitle)
        picTaskBuffer.CurrentX = TaskCenter(intWidth, strDrawText) + picTaskBuffer.CurrentX
        
        picTaskBuffer.CurrentY = picTaskBuffer.CurrentY + 3
        If intActive = i Then
            picTaskBuffer.CurrentX = picTaskBuffer.CurrentX + 1
        End If
        picTaskBuffer.Print strDrawText;
    Next i
    
    picTask.Picture = picTaskBuffer.Image
        
End Sub

Public Sub HandleCTCP(strNick As String, strData As String)
    strData = RightR(strData, 1)
    strData = LeftR(strData, 1)
    
    Dim strCom As String, strParam As String, inttemp As Integer, strTemp As String, strArgs() As String
    Dim dccinfo As DCC_INFO
    
    Seperate strData, " ", strCom, strParam
    
    Select Case LCase(strCom)
        Case "version"
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just requested your client version"
            CTCPReply strNick, "VERSION " & strVersionReply
        Case "ping"
            strTemp = RightOf(strData, " ")
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just pinged you"
            CTCPReply strNick, "PING " & strTemp
        Case "time"
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just requested the time on your machine."
            CTCPReply strNick, "TIME " & AscTime(CTime())
        Case "action"
            inttemp = GetQueryIndex(strNick)
            If inttemp = -1 Then Exit Sub
            PutText Queries(inttemp).DataIn, strColor & "06" & strNick & " " & strParam
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
                    dccinfo.Id = strNick & "." & strArgs(3)
                    
                    Dim bAccept As Boolean
                    bAccept = ShowAcceptDCC(dccinfo.Nick, dccinfo.File, dccinfo.Size)
                    
                    If bAccept = True Then
                        TimeOut 0.3
                        inttemp = NewDCCSend(dccinfo)
                        TimeOut 0.3
                        
                        If lngDCCStart = -1 Then 'DO NOT resume :P
                            DCCSends(inttemp).lngSentRcvd = 0
                            DCCSends(inttemp).lngFileStart = 0
                            DCCSends(inttemp).sock.Connect
                        Else                        'RESUME!!!
                            lngAccept = CLng(inttemp)
                            Client.SendData "PRIVMSG " & strNick & " :" & strAction & "DCC RESUME " & dccinfo.File & " " & dccinfo.Port & " " & lngDCCStart & strAction
                            
                            DCCSends(inttemp).lngSentRcvd = lngDCCStart
                            DCCSends(inttemp).lngFileStart = lngDCCStart
                            DCCSends(inttemp).sock.Close
                            DCCSends(inttemp).sock.LocalPort = dccinfo.Port
                            DCCSends(inttemp).sock.RemotePort = dccinfo.Port
                            DCCSends(inttemp).Caption = "DCC Receive - 00.00%"
                            DCCSends(inttemp).sock.Close
                        End If
                    Else
                        Client.SendData "NOTICE " & strNick & " :Your DCC request has been declined for " & dccinfo.File & " (" & Format(dccinfo.Size, "###,###,###,###") & " bytes)"
                    End If
                Case "accept"
                    DCCSends(lngAccept).StartResume DCCSends(lngAccept).sock.RemotePort
                    lngAccept = -1
                Case "chat"
                    Dim intRetX As Integer
                    inttemp = NewDCCChat(strNick, LongIPToIP(CLng(strArgs(2))), CLng(strArgs(3)))
                    If inttemp = -1 Then Exit Sub
                    
                    'put accept shit here...
                    intRetX = MsgBox("The user " & strNick & " has requested to open a DCC Chat session with you." & vbCrLf & "DCC Chat is not controlled by the server, therefore the other user may flood you if they wish to do so." & vbCrLf & "Would you like to open a DCC Chat session with this user?", vbYesNo Or vbQuestion)
                    If intRetX = vbNo Then
                        EchoActive "* DCC Chat request from " & strNick & " has been declined", 2
                        Exit Sub
                    End If
blah:
                    TimeOut 0.1
                    DoEvents
                    
                    PutData DCCChats(inttemp).DataIn, strColor & "02Attempting to make connection..."
                    DCCChats(inttemp).sock.Connect LongIPToIP(CLng(strArgs(2))), CLng(strArgs(3))
            End Select
            
    End Select
End Sub


Sub interpret(strData As String)
    Dim parsed As ParsedData, inttemp As Integer
    Dim i As Integer, StrChan As String, strTemp As String
    
    
    strData = Replace(strData, chr(10), "")
    ParseData strData, parsed
    'DoEvents
    If Len(parsed.strCommand) = 0 Then Exit Sub
    
    If parsed.strCommand <> "303" Then
        DebugWin.txtDataIn.SelStart = Len(DebugWin.txtDataIn.Text)
        DebugWin.txtDataIn.SelText = "  " & strData & vbCrLf
    End If
    
    '* I will put these in alphanumeric order one day, dont worry
    'MsgBox "~" & parsed.strCommand & "~"
    
    Select Case LCase(parsed.strCommand)
        Case "001"
            strMyNick = params(parsed, 1, 1)
            PutData Status.DataIn, "* " & params(parsed, 2, -1)
            strServer = parsed.strHost
            Status.Update
            TimeOut 3
            If strAutoJoin <> "<none>" Then SendData GetAlias("", "JOIN " & strAutoJoin)
            Exit Sub
        Case IRC_LOCALHOSTIRCD, IRC_SERVERCREATED, IRC_AVAILABLE, "004"
            PutData Status.DataIn, "* " & params(parsed, 2, -1)
            Exit Sub
        Case "ping"
            SendData "PONG :" & params(parsed, 1, 1)
            
            '* Show Ping? Pong! ??
            If Not bHidePing Then
                PutData Status.DataIn, strColor & "03Ping? Pong! [ " & params(parsed, 1, -1) & " ]"
            End If
            
            Exit Sub
        Case "join"
            If LCase(parsed.strNick) = LCase(strMyNick) Then
                Dim strNow As Double, strThen As Double, strDiff As Double
                strNow = GetTickCount()
                strThen = CDbl(ReadINI("lag", params(parsed, 1, -1), params(parsed, 1, -1)))
                
                inttemp = NewChannel(params(parsed, 1, -1))
                Client.SendData "MODE " & params(parsed, 1, -1)
                
                strResult = strNow - strThen
                
                '* If display lag...
                If bShowLag Then
                    TimeOut 0.5
                    PutData Channels(inttemp).DataIn, strColor & "03* Channel synched in " & Duration(CLng(strResult) / 1000) & " " & strResult Mod 1000 & " ms (LAGtime)"
                End If
            Else
                inttemp = GetChanIndex(params(parsed, 1, 1))
                If inttemp = -1 Then Exit Sub
                Channels(inttemp).AddNick parsed.strNick, parsed.strFullHost
                PutData Channels(inttemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " has joined " & strBold & Channels(inttemp).strName
            End If
            Exit Sub
        Case "privmsg"
            StrChan = params(parsed, 1, 1)
            If left(StrChan, 1) = "#" Or left(StrChan, 1) = "&" Then  'privmsg to channel
                inttemp = GetChanIndex(StrChan)
                If inttemp <> -1 Then Channels(inttemp).PutText parsed.strNick, params(parsed, 2, -1)                                  '
            ElseIf parsed.strNick = strMyNick Then
                If params(parsed, 2, 2) = strAction & "VERSION" & strAction Then    'version
                    'Client.SendData "CTCP REPLY " & strChan & " VERSION :jIRC for Windows9x"
                    Client.SendData "NOTICE " & parsed.strNick & " :VERSION " & strVersionReply & ""
                End If
                GoTo msg
            
            Else    'send to query window
msg:
                strTemp = params(parsed, 2, -1)
                If left(strTemp, 1) = strAction Then
                    HandleCTCP parsed.strNick, strTemp
                    Exit Sub
                End If
                
                If QueryExists(parsed.strNick) Then
                    inttemp = GetQueryIndex(parsed.strNick)
                    If inttemp = -1 Then Exit Sub
                    
                    If Queries(inttemp).strHost <> parsed.strFullHost Then
                        Queries(inttemp).strHost = RightOf(parsed.strFullHost, "!")
                        Queries(inttemp).lblHost = RightOf(parsed.strFullHost, "!")
                        
                    End If
                    Queries(inttemp).Caption = parsed.strNick
                    Queries(inttemp).strNick = parsed.strNick
                    Queries(inttemp).lblNick = parsed.strNick
                    Queries(inttemp).PutText parsed.strNick, strTemp
                Else
                    NewQuery parsed.strNick, parsed.strFullHost
                    inttemp = GetQueryIndex(parsed.strNick)
                    If inttemp = -1 Then Exit Sub
                    Queries(inttemp).Caption = parsed.strNick
                    Queries(inttemp).strNick = parsed.strNick
                    Queries(inttemp).lblNick = parsed.strNick
                    Queries(inttemp).PutText parsed.strNick, strTemp
                End If
            End If
            Exit Sub
        Case "nick"
            If parsed.strNick = strMyNick Then
                strMyNick = params(parsed, 1, 1)
                PutData Status.DataIn, strColor & "03Your nick is now " & strBold & strMyNick
                ChangeNick parsed.strNick, params(parsed, 1, -1)
                Status.Update
            Else
                ChangeNick parsed.strNick, params(parsed, 1, 1)
            End If
            Exit Sub
        Case "part"
            If parsed.strNick = strMyNick Then Exit Sub
            inttemp = GetChanIndex(parsed.strParams(1))
            If inttemp = -1 Then Exit Sub
            Channels(inttemp).RemoveNick parsed.strNick
            PutData Channels(inttemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " has left " & strBold & Channels(inttemp).strName
            If parsed.strNick = strMyNick Then Unload Channels(inttemp)
            Exit Sub
        Case "353" 'nick list!
            inttemp = GetChanIndex(parsed.strParams(3))
            If inttemp = -1 Then Exit Sub
            Dim strNicks() As String
            strNicks = Split(params(parsed, 4, -1), " ")
            For i = LBound(strNicks) To UBound(strNicks)
                Channels(inttemp).AddNick strNicks(i)
            Next i
            Exit Sub
        Case "mode"     'set mode
            inttemp = GetChanIndex(params(parsed, 1, 1))
            If inttemp = -1 And params(parsed, 1, 1) <> strMyNick Then Exit Sub
            strTemp = parsed.strNick
            If strTemp = "" Then strTemp = parsed.strFullHost
            
            If inttemp <> -1 Then PutData Channels(inttemp).DataIn, strColor & "03" & strBold & strTemp & strBold & " sets mode: " & params(parsed, 2, -1)
            If params(parsed, 1, 1) = strMyNick Then PutData Status.DataIn, strColor & "03" & strBold & strTemp & strBold & " sets mode: " & params(parsed, 2, -1)
            
            ParseMode params(parsed, 1, 1), params(parsed, 2, -1)
            Exit Sub
        Case "quit"     'quit
            NickQuit parsed.strNick, params(parsed, 1, -1)
            Exit Sub
        Case "kick"     'kick
            inttemp = GetChanIndex(params(parsed, 1, 1))
            If inttemp = -1 Then Exit Sub
            PutData Channels(inttemp).DataIn, strColor & "03" & strBold & params(parsed, 2, 2) & strBold & " was kicked from " & strBold & params(parsed, 1, 1) & strBold & " by " & strBold & parsed.strNick & strBold & " [ " & params(parsed, 3, -1) & " ]"
            Channels(inttemp).RemoveNick params(parsed, 2, 2)
            
            '* If user, close channel
            If params(parsed, 2, 2) = strMyNick Then
                '* close channel w/o sending PART command
                Channels(inttemp).Tag = "NOPART"
                '* unload it
                Unload Channels(inttemp)
                PutData Status.DataIn, strColor & "03" & "You were kicked from " & strBold & params(parsed, 1, 1) & strBold & " by " & strBold & parsed.strNick & strBold & " [ " & params(parsed, 3, -1) & " ]"
                
                '* Rejoin when kicked?
                If bRejoinOnKick Then
                    TimeOut 0.1
                    Client.SendData "JOIN " & params(parsed, 1, 1)
                End If
            End If
            
            Exit Sub
        Case "332"  'topic!
            inttemp = GetChanIndex(params(parsed, 2, 2))
            If inttemp = -1 Then Exit Sub
            Channels(inttemp).rtbTopic.Text = ""
            PutData Channels(inttemp).rtbTopic, params(parsed, 3, -1)
            Channels(inttemp).rtbTopic.SelStart = 0
            Channels(inttemp).rtbTopic.SelLength = 1
            Channels(inttemp).rtbTopic.SelText = ""
            PutData Channels(inttemp).DataIn, strColor & "03Topic is """ & strColor & params(parsed, 3, -1) & strColor & "03"""
            Channels(inttemp).rtbTopic.SelStart = 0
            Channels(inttemp).rtbTopic.Tag = "locked"
            Channels(inttemp).strTopic = params(parsed, 3, -1)
            Exit Sub
        Case "topic"    'change in topic!
            inttemp = GetChanIndex(params(parsed, 1, 1))
            If inttemp = -1 Then Exit Sub
            Channels(inttemp).rtbTopic.Text = ""
            PutData Channels(inttemp).rtbTopic, params(parsed, 2, -1)
            Channels(inttemp).rtbTopic.SelStart = 0
            Channels(inttemp).rtbTopic.SelLength = 1
            Channels(inttemp).rtbTopic.SelText = ""
            Channels(inttemp).strTopic = params(parsed, 3, -1)
            PutData Channels(inttemp).DataIn, strColor & "03Topic changed by " & strBold & parsed.strNick & strBold & " : " & params(parsed, 2, -1)
            Exit Sub
        Case "333"  'topic on param2 set by param3, on param4
            inttemp = GetChanIndex(params(parsed, 2, 2))
            If inttemp = -1 Then Exit Sub
            PutData Channels(inttemp).DataIn, strColor & "03Topic set by " & strBold & params(parsed, 3, 3) & strBold & " on " & strBold & AscTime(params(parsed, 4, 4)) & strBold
            '* gotta add when it was SET!!
            Exit Sub
        Case "366"  'end of names list
            Exit Sub
        Case "324"  'set channel modes
            ParseMode params(parsed, 2, 2), params(parsed, 3, -1)
            Exit Sub
        Case "notice"
            'On Error Resume Next
            
            Seperate params(parsed, 2, -1), " ", StrChan, strTemp 'strChan is actual the type
            
            '* CTCP Replies
            If left(params(parsed, 2, 2), 1) = strAction Then
                PutCTCPReply parsed.strNick, RightR(StrChan, 1), LeftR(strTemp, 1)
                Exit Sub
            End If
            
            '* Actual notices
            If parsed.strNick = "" Then
                PutData Status.DataIn, strColor & "05" & params(parsed, 2, -1)
            Else
                EchoActive strColor & "05[notice]" & strColor & " <- " & strBold & parsed.strNick & strBold & ": " & chr(9) & params(parsed, 2, -1)
                'EchoActive strColor & "05" & strBold & "NOTICE" & strBold & strColor & " " & strBold & parsed.strNick & strBold & ":" & Chr(9) & params(parsed, 2, -1)
                Exit Sub
                
                If Client.ActiveForm.hWnd = BuddyList.hWnd Then
                    PutData Status.DataIn, strColor & "05" & strBold & "NOTICE" & strBold & strColor & " " & strBold & parsed.strNick & strBold & ":" & chr(9) & params(parsed, 2, -1)
                Else
                    PutData Client.ActiveForm.DataIn, strColor & "05" & strBold & "NOTICE" & strBold & strColor & " " & strBold & parsed.strNick & strBold & ":" & chr(9) & params(parsed, 2, -1)
                End If
            End If
            
            Exit Sub
        Case "433"  'nick name already in use
            If params(parsed, 2, 2) = strMyNick Then
                strTemp = strMyNick
                strMyNick = strOtherNick
                strOtherNick = strMyNick
                Client.SendData "NICK " & strMyNick
            End If
            PutData Status.DataIn, strColor & "04* " & params(parsed, 2, -1)
            Exit Sub
        Case "372"  'MOTD
            PutData Status.DataIn, params(parsed, 2, -1)
            Exit Sub
        Case "375"  'start of MOTD
            PutData Status.DataIn, strColor & "02" & params(parsed, 2, -1)
            Exit Sub
        Case "376"  'end of MOTD
            If bInvisible Then Client.SendData "MODE " & strMyNick & " +i"
            Exit Sub
        Case "251", "252", "253", "254", "255", _
            "265", "266" 'server info, users, ops, channels, clients
            PutData Status.DataIn, strColor & "06" & params(parsed, 2, -1)
            Exit Sub
        Case "303"  'users on!
            BuddyList.AddUsers params(parsed, 2, -1)
            Exit Sub
        Case "329"  'date created for channel, $2 = channel, 3 = when
            inttemp = GetChanIndex(params(parsed, 2, 2))
            If inttemp = -1 Then Exit Sub
            PutData Channels(inttemp).DataIn, strColor & "03Channel created on " & strBold & AscTime(params(parsed, 3, 3))
            Exit Sub
        Case "484", "482", "461", "412", "403", "421", "401", "481", "402", "451", "404" 'error msgs
        
            EchoActive "* " & params(parsed, 2, -1), 4
            Exit Sub
        Case "321"  'blah channel list
            'ChannelsList.lvChannels.ListItems.Clear
            ChannelsList.Visible = True
            Exit Sub
        Case "322"          'channel list! ACK!!
            ChannelsList.AddChannel params(parsed, 2, 2), params(parsed, 3, 3), params(parsed, 4, -1)
            ChannelsList.Caption = "Channels List : " & ChannelsList.lvChannels.ListItems.Count & " channels"
            Exit Sub
        Case "323"  'end of channe list!
            'ChannelsList.bNeedClear = True
            'ChannelsList.lvChannels.ListItems.Clear
            Exit Sub
        Case "328"
            inttemp = GetChanIndex(params(parsed, 2, 2))
            If inttemp = -1 Then Exit Sub
            PutData Channels(inttemp).DataIn, strColor & "03Channel URL" & strColor & strUnderline & ":" & strUnderline & " " & strColor & "12" & params(parsed, 3, -1) & strColor
            Exit Sub
        Case "301"      'WHOIS! away
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is away: " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " is away: " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is away: " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "307"  'nick ident
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "311"  'whois, ident, host, etc
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, " "
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
                Else
                    PutData Queries(inttemp).DataIn, " "
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " is " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
                End If
            Else
                PutData Status.DataIn, " "
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
            End If
            Exit Sub
        Case "312"  'server
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " using: " & params(parsed, 3, 3) & " ( " & params(parsed, 4, -1) & " )"
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " using: " & params(parsed, 3, 3) & " ( " & params(parsed, 4, -1) & " )"
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is using: " & params(parsed, 3, 3) & " ( " & params(parsed, 4, -1) & " )"
            End If
            Exit Sub
        Case "313"  'is an irc operator
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "314"  'whowas info, ident, host, etc
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, " "
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " was " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
                Else
                    PutData Queries(inttemp).DataIn, " "
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " was " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
                End If
            Else
                PutData Status.DataIn, " "
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " was " & params(parsed, 3, 3) & "@" & params(parsed, 4, -1)
            End If
            Exit Sub
        Case "317"
            Dim strPut As String, strPut2 As String
            strTemp = params(parsed, 2, 2)
            strPut = strColor & "05" & strBold & strTemp & strBold & " has been idle for : " & Duration(params(parsed, 3, 3))
            strPut2 = strColor & "05 " & strBold & strTemp & strBold & " signed on at : " & AscTime(CLng(params(parsed, 4, 4)))
            strPut = strPut & vbCrLf & strPut2
            
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strPut 'strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, 3) & " seconds idle / " & AscTime(CLng(params(parsed, 4, 4))) & " sign-on time"
                Else
                    PutData Queries(inttemp).DataIn, strPut 'strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, 3) & " seconds idle / " & AscTime(CLng(params(parsed, 4, 4))) & " sign-on time"
                End If
            Else
                PutData Status.DataIn, strPut 'strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, 3) & " seconds idle / " & AscTime(CLng(params(parsed, 4, 4))) & " sign-on time"
            End If
            Exit Sub
        Case "318"
            Exit Sub
        Case "319"
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is on: " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " is on: " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " is on: " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "320"
            strTemp = params(parsed, 2, 2)
            Seperate strTemp, " ", StrChan, strTemp
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(StrChan)
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & StrChan & strBold & " " & strTemp
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & StrChan & strBold & " " & strTemp
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & StrChan & strBold & " " & strTemp
            End If
            Exit Sub
        Case "305", "306"   '305=no longer away, 306=marked as away
            PutData Status.DataIn, strColor & "05" & params(parsed, 2, -1)
            Exit Sub
        Case "369"  'end of whowas
            Exit Sub
        Case "391"  'time
            PutData Status.DataIn, "Time: " & params(parsed, 5, -1)
            Exit Sub
        Case "462"
            PutData Status.DataIn, strColor & "04" & params(parsed, 2, -1)
            Exit Sub
        Case "501"      'unknown mode flag
            PutData Status.DataIn, strColor & "04" & params(parsed, 2, -1)
            Exit Sub
        Case "472"      'param1 is unknown mode char to me
            PutData Status.DataIn, strColor & "04" & strBold & params(parsed, 2, 2) & strBold & " " & params(parsed, 3, -1)
            Exit Sub
        Case "406"  'there was no such nick
            EchoActive strBold & params(parsed, 2, 2) & strBold & " " & params(parsed, 3, -1), 4
            Exit Sub
        Case "315"  'end of who list
            Exit Sub
        Case "438"
            PutData Status.DataIn, strColor & "04" & params(parsed, 1, 1) & " -> " & params(parsed, 2, 2) & ": " & params(parsed, 3, -1)
            Exit Sub
        Case "405"  'cannot join param2, param3
            PutData Status.DataIn, strColor & "04Cannot join " & params(parsed, 2, 2) & " ( " & params(parsed, 3, -1) & " )"
            Exit Sub
        Case "471", "473", "474", "475"   'cannot join channel, 471=+l, 473=+i, 475=+k, 474=+b
            PutData Status.DataIn, strColor & "04Cannot join " & params(parsed, 2, 2) & ": " & params(parsed, 3, -1)
            Exit Sub
        Case "263"
            PutData Status.DataIn, "* " & params(parsed, 2, -1)
            Exit Sub
        Case "617"
            PutData Status.DataIn, strColor & "05* " & params(parsed, 2, -1)
            Exit Sub
        Case "error"
            EchoActive "* " & params(parsed, 1, -1), 4
            Exit Sub
        Case "331"  'no topic set
            inttemp = GetChanIndex(params(parsed, 2, 2))
            If inttemp = -1 Then
                PutData Status.DataIn, strColor & "03(" & params(parsed, 2, 2) & ") " & params(parsed, 3, -1)
            Else
                PutData Channels(inttemp).DataIn, strColor & "03(" & params(parsed, 2, 2) & ") " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "335"  'bot on server
            strTemp = params(parsed, 2, 2)
            If bWhoisInQuery Then
                inttemp = GetQueryIndex(strTemp)
                'MsgBox inttemp
                If inttemp = -1 Then
                    PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                Else
                    PutData Queries(inttemp).DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
                End If
            Else
                PutData Status.DataIn, strColor & "05" & strBold & strTemp & strBold & " " & params(parsed, 3, -1)
            End If
            Exit Sub
        Case "341"  'invited param2 to param3
            EchoActive "* Invited " & strBold & params(parsed, 2, 2) & strBold & " to " & strBold & params(parsed, 3, 3) & strBold, 3
            Exit Sub
        Case "479", "442"
            '479 = channel name contains illegal chars
            '442 = you're not on that channel
            EchoActive "* " & strBold & params(parsed, 2, 2) & strBold & " " & params(parsed, 3, -1), 4
            Exit Sub
        Case "465"  'you have been k-lined
            EchoActive "* " & params(parsed, 2, -1), 2
        Case "491"  'no olines for your host
            EchoActive "* " & params(parsed, 2, -1), 5
            Exit Sub
        Case "502"  'cannot set mode for others
            EchoActive "* " & params(parsed, 2, -1), 3
            Exit Sub
        Case "512"  'no such GLine
            EchoActive "* " & strBold & params(parsed, 2, 2) & strBold & " " & params(parsed, 3, -1), 5
            Exit Sub
        Case "invite"   'invite
            EchoActive "* " & strBold & parsed.strNick & strBold & " has invited you to " & strBold & params(parsed, 2, 2), 3
            Exit Sub
        Case "371"  'info
            PutData Status.DataIn, strColor & "06* " & params(parsed, 2, -1)
            Exit Sub
        Case "374"  'end of info
            Exit Sub
        Case "367"
            If bGettingChanInfo Then
                ChannelInfo.AddBEI params(parsed, 3, 3), params(parsed, 4, 4), params(parsed, 5, 5)
            End If
            Exit Sub
        Case "368"  'end of ban list
            If ChannelInfo.Visible = False Then ChannelInfo.Show vbModal
            bGettingChanInfo = False
            Exit Sub
        Case "kill"
            If params(parsed, 1, 1) = strMyNick Then
                EchoActive strBold & "YOU" & strBold & " were KILLED [ " & params(parsed, 2, -1) & " ]", 2
            Else
                EchoActive strBold & params(parsed, 1, 1) & strBold & " was KILLED [ " & params(parsed, 2, -1) & " ]", 2
            End If
            Exit Sub
    End Select
    PutData Status.DataIn, "*** " & strBold & parsed.strCommand & strBold & " " & strBold & strTemp & strBold & " " & params(parsed, 1, -1) ' & " [" & parsed.strFullHost & "]"
End Sub


Sub SendData(strData As String)
    On Error Resume Next
    If sock.State = 0 Then Exit Sub
    sock.SendData strData & chr(10)
    If DebugWin.Visible Then
        If strData Like "ISON *" Then Else DebugWin.txtDataIn = DebugWin.txtDataIn & ">> OUTGOING DATA >> " & vbCrLf & strData & vbCrLf
        DebugWin.txtDataIn.SelStart = Len(DebugWin.txtDataIn)
    End If
    
End Sub




Sub SetClientPos()
    Dim strPos As String, strCPos As String, strLst() As String
    strCPos = "-1,-1,-1,-1"
    strPos = GetINI(winINI, "pos", "!client", strCPos)
    strLst = Split(strPos, ",")
    If UBound(strLst) <> 3 Then Exit Sub
    
    
    If CInt(strLst(0)) = -1 Then Exit Sub
    If CInt(strLst(1)) = -1 Then Exit Sub
    If CInt(strLst(2)) = -1 Then Exit Sub
    If CInt(strLst(3)) = -1 Then Exit Sub
    
    Client.Move CInt(strLst(0)), CInt(strLst(1)), CInt(strLst(2)), CInt(strLst(3))
    
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    i = NewDCCChat("he", "host", 10)
End Sub

Sub ToggleMenu(bShow As Boolean)
    mnu_File.Visible = bShow
    mnu_view.Visible = bShow
    mnu_window.Visible = bShow
    mnu_Help.Visible = bShow
    mnu_commands.Visible = bShow
    mnu_Edit.Visible = bShow
End Sub

Private Sub cScript_Error()
    EchoActive "* An Error has occured on Line #" & cScript.Error.Line & " : " & cScript.Error.description, 4
End Sub

Private Sub IDENT_ConnectionRequest(ByVal requestID As Long)
    IDENT.Close
    IDENT.Accept requestID
    IDENT.SendData IDENT.LocalPort & ", " & IDENT.RemotePort & " : USERID : UNIX : " & strMyIdent & vbCrLf
    
    Dim i As Integer
    For i = 1 To 500
        DoEvents
    Next i
    
    IDENT.Close
End Sub

Private Sub IDENT_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String
    IDENT.GetData dat, vbString
    
    If dat Like "*, *" Then
        dat = LeftR(dat, 2)
        PutData Status.DataIn, "*** IDENT : " & dat
        dat = dat & " : USERID : UNIX : " & strMyIdent
        On Error Resume Next
        IDENT.SendData dat
        PutData Status.DataIn, "*** IDENT reply : " & dat
        Dim i As Integer
        IDENT.Close
    End If
End Sub

Private Sub IDENT_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData Status.DataIn, chr(Color) & "04IDENT Error " & strColor & description
End Sub

Private Sub MDIForm_Load()
                
                
    'strVersionReply = "projectIRC " & App.Major & "." & App.Minor & App.Revision
    strVersionReply = "MicrowaveIRC 1.2 for Microwaves"
                
    intCTries = 1
                
    picToolMain.Line (3, 4)-(5, picToolMain.ScaleHeight - 5), &H80000014, B
    picToolMain.Line (5, 4)-(5, picToolMain.ScaleHeight - 5), &H80000010
    picToolMain.Line (4, picToolMain.ScaleHeight - 5)-(6, picToolMain.ScaleHeight - 5), &H80000010
    
    picToolMain.Line (6, 4)-(8, picToolMain.ScaleHeight - 5), &H80000014, B
    picToolMain.Line (8, 4)-(8, picToolMain.ScaleHeight - 5), &H80000010
    picToolMain.Line (7, picToolMain.ScaleHeight - 5)-(9, picToolMain.ScaleHeight - 5), &H80000010
                
                    
    IFTiface = IFT_FANCY        'Set fancy interface, for now
    'intButtonWidth = 100        'button fixed width (now in ini, can be set)
    sngLastX = 1                'last position of x, used for taskbar
    DoEvents
    Me.Visible = True
    
    Load LoadSettings           'load dialog which displays "Loading settings..."
    LoadSettings.Visible = True
    DoEvents
    StayOnTop LoadSettings, True
        
    DrawToolbar                 'draw toolbar
    
    '* INI stuff
    path = App.path
    If Right(App.path, 1) <> "\" Then path = path & "\"
    
    INI = path & strUserProfile & "-settings.ini"
    winINI = path & strUserProfile & "-windows.ini"
    
    SetClientPos
    '/if doesnt exist, create
    If Not FileExists(path & "settings.ini") Then
        Open INI For Output As #1
            Print #1, ""
        Close #1
    End If
    
    '/* Connection settings
    strServer = ReadINI("connect", "server", "irc.otherside.com")
    strMyNick = ReadINI("connect", "nick", "pIRCu")
    strOtherNick = ReadINI("connect", "altnick", "OtherNick")
    strFullName = ReadINI("connect", "fullname", "projectIRC user")
    strMyIdent = ReadINI("connect", "ident", "projectIRC")
    lngPort = CLng(ReadINI("connect", "port", "6667"))
    bConOnLoad = CBool(ReadINI("connect", "connonload", "false"))
    bReconnect = CBool(ReadINI("connect", "reconnect", "true"))
    bInvisible = CBool(ReadINI("connect", "invisible", "true"))
    bRetry = CBool(ReadINI("connect", "retry", "true"))
    intRetry = CInt(ReadINI("connect", "retrynum", "99"))
    strAutoJoin = ReadINI("connect", "autojoin", "#projectIRC")
    
    '/* Display settings
    lngBackColor = CLng(ReadINI("display", "backcolor", CStr(RGB(255, 255, 255))))
    lngForeColor = CLng(ReadINI("display", "forecolor", CStr(RGB(0, 0, 0))))
    lngLeftColor = CLng(ReadINI("display", "leftcolor", CStr(&HA27E66)))
    lngRightColor = CLng(ReadINI("display", "rightcolor", CStr(&H8000000F)))
    strFontName = ReadINI("display", "fontname", "terminal")
    intFontSize = CInt(ReadINI("display", "fontsize", CStr(8)))
    lngClientBack = CLng(ReadINI("display", "clientbg", CStr(&H8000000C)))
    Client.BackColor = lngClientBack
    strBGImage = ReadINI("display", "bgimage", "")
    bBGImage = CBool(ReadINI("display", "bgimageon", "false"))
    bTileImg = CBool(ReadINI("display", "tileimg", "true"))
    
    '* Window/Interface Settings
    bFlatButtons = CBool(ReadINI("windows", "flatbuttons", "true"))
    bStretchButtons = CBool(ReadINI("windows", "stretch", "true"))
    intButtonWidth = CInt(ReadINI("windows", "buttonwidth", "100"))
    bMinToSystray = CBool(ReadINI("windows", "mintosystray", "true"))
    bAlwaysShowST = CBool(ReadINI("windows", "alwaysshowst", "false"))
    intTranslucency = CInt(ReadINI("windows", "translucency", "0"))
    'MsgBox intTranslucency
    
    '* IRC Settings
    bWhoisInQuery = CBool(ReadINI("irc", "whoisquery", "false"))
    MAX_TEXT_HISTORY = CInt(ReadINI("irc", "maxtexthistory", "30"))
    bAnnounce = CBool(ReadINI("irc", "announceaway", "true"))
    bRejoinOnKick = CBool(ReadINI("irc", "rejoinonkick", "true"))
    bHidePing = CBool(ReadINI("irc", "hidepong", "true"))
    bNickComplete = CBool(ReadINI("irc", "nickcomplete", "true"))
    strQuitMsg = ReadINI("irc", "quitmsg", "using projectIRC, closed")
    bShowLag = CBool(ReadINI("irc", "showlag", "true"))
    
    '* Systray
    SysTray.Initialize Me.hWnd, Me.Icon, "projectIRC"
    If bAlwaysShowST Then SysTray.ShowIcon
    
    '* Other stuff..
    bDock = CBool(ReadINI("interface", "dock", "true"))
    
    DoEvents
    
    

    '* too freaking SLOW
    'If intTranslucency <> 0 Then Call trans.fSetTranslucency(Client.hWnd, (100 - intTranslucency) * 2.5)
        
    Load Status                 'load status dialog
    Status.SetFocus
    
    '* Draw BG
    If bBGImage Then
        On Error Resume Next
        '*normal
        Client.Picture = LoadPicture(strBGImage)
        picBackTemp.Picture = Client.Picture
        
        '*tile
        If bTileImg Then
            Dim X As Integer, Y As Integer
            Dim i As Double, j As Double
            
            i = Screen.Width / picBackTemp.ScaleWidth + 1
            j = Screen.Height / picBackTemp.ScaleHeight + 1
            
            picBack.Height = Screen.Height / 15
            picBack.Width = Screen.Width / 15
            
            For X = 0 To i
                For Y = 0 To j
                    BitBlt picBack.hDC, X * picBackTemp.Width, Y * picBackTemp.Height, picBackTemp.Width, picBackTemp.Height, picBackTemp.hDC, 0, 0, SRCCOPY
                Next Y
            Next X
            
            Client.Picture = picBack.Image
            Client.Visible = False
            Client.Visible = True
        End If
        'picBack.Picture = LoadPicture("")
    End If
    
    '* UnDock toolbar?
    If bDock = False Then
        Load toolbardock
        toolbardock.Show
        SetParent picToolMain.hWnd, toolbardock.hWnd
        bDock = False
        picToolMain.Tag = ""
    End If
    
    '* Initialize script
    InitializeScript
    
    '* Load aliases
    LoadAliases
    
    Unload LoadSettings
    
    
End Sub

 
Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'When the callback message of CSystrayIcon is WM_MOUSEMOVE,
    'the X of Form_MouseMove is used to see what happen to the
    'icon in the systray.
    Dim msgCallBackMessage As Long
    
    'To be able to compare the callback value to the window message,
    'we must divide X by Screen.TwipsPerPixelX. That represent the
    'horizontal number of twips in the screen. (1 pixel ~= 15 twips)
    msgCallBackMessage = X / Screen.TwipsPerPixelX
     
    Select Case msgCallBackMessage
        Case WM_MOUSEMOVE
          
        'Case WM_LBUTTONDOWN
          
        'Case WM_LBUTTONUP
          
        Case WM_LBUTTONDBLCLK
            'SysTray.HideIcon
            Me.Visible = True
            Me.WindowState = vbNormal
            Me.Visible = True
            If Not bAlwaysShowST Then SysTray.HideIcon
        Case WM_RBUTTONDOWN
          
        Case WM_RBUTTONUP
            If Me.WindowState = vbMinimized Then PopupMenu mnu_Systray, , , , mnu_Systray_Restore
        Case WM_RBUTTONDBLCLK
          
        Case WM_MBUTTONDOWN
          
        Case WM_MBUTTONUP
          
        Case WM_MBUTTONDBLCLK
              
    End Select
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Unload BuddyList
    Unload Status
    
    toolbardock.Tag = "dontdock"
    Unload toolbardock
    
    Unload DebugWin
    Client.SendData "QUIT :" & strQuitMsg
    
    Dim i As Integer
    
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then Unload Channels(i)
    Next
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then Unload Queries(i)
    Next i
    
    For i = 0 To 15
        DoEvents
    Next i
    
    Client.Picture = LoadPicture("")
    SysTray.HideIcon
    Unload Me
    'Cancel = 0
    End
End Sub

Private Sub MDIForm_Resize()

    If Me.WindowState = vbMinimized Then
        If bMinToSystray Then
            SysTray.ShowIcon
            Me.Visible = False
        End If
        Exit Sub
    End If
'    lnTB2.X1 = (Client.Width / 15) - 9
'    lnTB2.X2 = lnTB2.X1

'    CoolBar1.Width = (Me.ScaleWidth / 15) + 1
    SetWinFocus intActive
    DrawToolbar
    
    If bSBL = True Then
        DoEvents
        If BuddyList.Visible = True Then BuddyList.Move Client.Width - BuddyList.Width - 180, 0
    End If
    
End Sub

Private Sub MDIForm_Terminate()
'    Client.SendData "QUIT :Client closed, using projectIRC"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
'    Call CoolMenuObj.Install(0&)
'    Set CoolMenuObj = Nothing
    TimeOut 0.01
    DoEvents
    
    SysTray.HideIcon
    End
End Sub

Private Sub mnu_commands_join_Click()
    Dim strInput As String
    strInput = InputX("Enter channel(s) to join, seperated each by a comma" & vbCrLf & "(EX: ""#irc"" or ""#irc, #mIRC"")." & vbCrLf & vbCrLf & "Also if you don't prefix # or &&, I will for you (EX: ""irc,help"")", "Join channels")
    If strInput = chr(8) Then Exit Sub
    strInput = GetAlias("", "JOIN " & strInput)
End Sub

Private Sub mnu_commands_noticeuser_Click()
    Dim strMsg As String, strSend As String
    strNck = InputX("Enter the nickname of the user you would like to notice", "Send Notice to user")
    If strNck = chr(8) Then Exit Sub
    strMsg = InputX("Enter the Message you would like to send to " & strNck & ".", "Notice Message")
    If strNck = chr(8) Then Exit Sub
    strSend = GetAlias("Status", "Notice " & strNck & " " & strMsg)
    Client.SendData strSend

End Sub

Private Sub mnu_commands_query_Click()
    Dim strMsg As String, strSend As String
    strNck = InputX("Enter the nickname of the user you would like to query", "Query user")
    If strNck = chr(8) Then Exit Sub
    strMsg = InputX("Enter the Message you would like to send to " & strNck & ".", "Query User Message")
    If strNck = chr(8) Then Exit Sub
    strSend = GetAlias("Status", "QUERY " & strNck & " " & strMsg)
    Client.SendData strSend

End Sub

Private Sub mnu_Edit_Copy_Click()
On Error Resume Next
    With Client.ActiveForm.ActiveControl
        Clipboard.SetText .SelText
    End With
End Sub

Private Sub mnu_Edit_cut_Click()
    On Error Resume Next
    If Client.ActiveForm.ActiveControl.Name = "DataIn" Then Exit Sub
    With Client.ActiveForm.ActiveControl
        Clipboard.SetText .SelText
        .SelText = ""
    End With
End Sub

Private Sub mnu_Edit_Delete_Click()
If Client.ActiveForm.ActiveControl.Name = "DataIn" Then Exit Sub
    With Client.ActiveForm.ActiveControl
        .SelText = ""
    End With
End Sub

Private Sub mnu_Edit_Lcase_Click()
    On Error Resume Next
If Client.ActiveForm.ActiveControl.Name = "DataIn" Then Exit Sub
    With Client.ActiveForm.ActiveControl
        Dim intStart As Long, intLen As Long
        intStart = .SelStart
        intLen = .SelLength
        
        .SelText = LCase(.SelText)
        
        .SelStart = intStart
        .SelLength = intLen
    End With
End Sub

Private Sub mnu_Edit_Paste_Click()
On Error Resume Next
If Client.ActiveForm.ActiveControl.Name = "DataIn" Then Exit Sub
    With Client.ActiveForm.ActiveControl
        .SelText = Clipboard.GetText
    End With
End Sub

Private Sub mnu_Edit_selectall_Click()
On Error Resume Next
    With Client.ActiveForm.ActiveControl
        .SelStart = 0
        .SelLength = Len(.Text)
        '.SelStart = .SelLength
    End With
End Sub


Private Sub mnu_Edit_UCase_Click()
    On Error Resume Next
    If Client.ActiveForm.ActiveControl.Name = "DataIn" Then Exit Sub
    With Client.ActiveForm.ActiveControl
Dim intStart As Long, intLen As Long
        intStart = .SelStart
        intLen = .SelLength
        
        .SelText = UCase(.SelText)
        
        .SelStart = intStart
        .SelLength = intLen
    End With
End Sub


Private Sub mnu_Edit_Undo_Click()
    On Error Resume Next
    SendMessage Client.ActiveForm.ActiveControl.hWnd, EM_UNDO, 0&, 0
End Sub

Sub mnu_File_Connect_Click()
    Select Case mnu_File_Connect.Caption
        Case "&Connect"
            '* Connect
            sock.Close
            mnu_File_Connect.Caption = "&Cancel"
            sock.RemoteHost = strServer
            sock.RemotePort = lngPort
            IDENT.Close
            On Error Resume Next
            IDENT.Listen
            sock.Connect
            PutData Status.DataIn, strColor & "02Connecting to " & strBold & strServer & strBold & " port " & strBold & lngPort
        Case "&Cancel"
            '* Cancel
            IDENT.Close
            sock.Close
            mnu_File_Connect.Caption = "&Connect"
            PutData Status.DataIn, strColor & "05Connection attempt cancelled"
    End Select
End Sub


Sub mnu_File_Disconnect_Click()
    Client.SendData "QUIT :" & strQuitMsg
    Dim i As Integer
    For i = 1 To 300
        DoEvents
    Next i
    TimeOut 0.2
    mnu_File_Connect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    mnu_File_Disconnect.Enabled = False
    PutData Status.DataIn, strColor & "05Disconnected from " & sock.RemoteHost
    sock.Close
    ClearModes
    IDENT.Close
    Status.lblServer = "not connected"
    Status.Update
    BuddyList.lstNicks.Clear
    
    intCTries = 1
End Sub


Private Sub mnu_File_minSystray_Click()
    SysTray.ShowIcon
    Me.Visible = False
End Sub

Private Sub mnu_File_Options_Click()
    Options.Show ' 1
End Sub

Private Sub mnu_File_Quit_Click()
    Dim intMsg  As Integer
    intMsg = MsgBox("Would you really to like to exit projectIRC?", vbQuestion Or vbYesNo)
    If intMsg = vbNo Then Exit Sub
    
    Client.SendData "QUIT :Using projectIRC, closed"
    IDENT.Close
    sock.Close
    Dim i As Integer
    For i = 1 To 1000
        DoEvents
    Next i
    Unload Me
End Sub

Private Sub mnu_File_ScriptAliases_Click()
    AliasesEditor.Show vbModal
End Sub

Private Sub mnu_Help_About_Click()
    About.Show vbModal
End Sub

Private Sub mnu_nicks_Ban_Click()
    Dim StrChan As String, strNick As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    Client.SendData "MODE " & StrChan & " +b " & strNick & "!*@*"
End Sub

Private Sub mnu_nicks_CTCP_PING_Click()
Dim StrChan As String, strNick As String, strSend As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    strSend = GetAlias(StrChan, "PING " & strNick)
    Client.SendData strSend
End Sub

Private Sub mnu_nicks_CTCP_TIME_Click()
Dim StrChan As String, strNick As String, strSend As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    strSend = GetAlias(StrChan, "CTCP " & strNick & " TIME")
    Client.SendData strSend
End Sub


Private Sub mnu_nicks_CTCP_VERSION_Click()
    Dim StrChan As String, strNick As String, strSend As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    strSend = GetAlias(StrChan, "VERSION " & strNick)
    Client.SendData strSend
End Sub


Private Sub mnu_nicks_dccchat_Click()
    Dim StrChan As String, strNick As String, strSend As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    strSend = GetAlias(StrChan, "DCCCHAT " & strNick)
    Client.SendData strSend
End Sub

Private Sub mnu_nicks_dccsend_Click()
    Dim StrChan As String, strNick As String, strSend As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    strSend = GetAlias(StrChan, "DCCSEND " & strNick)
    Client.SendData strSend
End Sub


Sub mnu_nicks_halfop_Click()
    Dim StrChan As String, strNick As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    If mnu_nicks_halfop.Caption = "&HalfOp" Then
        Client.SendData "MODE " & StrChan & " +h " & strNick
    Else
        Client.SendData "MODE " & StrChan & " -h " & strNick
    End If
        
End Sub

Private Sub mnu_nicks_Kick_Click()
    Dim StrChan As String, strNick As String, strReason As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    strReason = InputX("Why would you like to kick the user '" & strNick & "' from " & StrChan & "?", "Kick User")
    If strReason = chr(8) Then Exit Sub
    
    Client.SendData "KICK " & StrChan & " " & strNick & " :" & strReason
        
End Sub

 Sub mnu_nicks_op_Click()
    Dim StrChan As String, strNick As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    If mnu_nicks_op.Caption = "&Op" Then
        Client.SendData "MODE " & StrChan & " +o " & strNick
    Else
        Client.SendData "MODE " & StrChan & " -o " & strNick
    End If
End Sub

Sub mnu_nicks_voice_Click()
    Dim StrChan As String, strNick As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    If mnu_nicks_voice.Caption = "&Voice" Then
        Client.SendData "MODE " & StrChan & " +v " & strNick
    Else
        Client.SendData "MODE " & StrChan & " -v " & strNick
    End If
End Sub

Private Sub mnu_nicks_WhoIs_Click()
Dim StrChan As String, strNick As String, strSend As String
    StrChan = Client.ActiveForm.strName
    strNick = RealNick(Client.ActiveForm.lstNicks.List(Client.ActiveForm.lstNicks.ListIndex))
    
    strSend = GetAlias(StrChan, "WHOIS " & strNick)
    Client.SendData strSend
End Sub

Private Sub mnu_Systray_Quit_Click()
    Call mnu_File_Quit_Click
End Sub

Private Sub mnu_Systray_Restore_Click()
    Me.Visible = True
    Me.WindowState = vbNormal
    Me.Visible = True
    If Not bAlwaysShowST Then SysTray.HideIcon
End Sub

Private Sub mnu_Tile_Vertically_Click()
    Client.Arrange vbTileVertical
End Sub

Private Sub mnu_view_3darea_Click()
    mnu_view_flatarea.Checked = False
    mnu_view_3darea.Checked = True
    Client.Appearance = 1
End Sub

Private Sub mnu_View_BuddyList_Click()
    With mnu_View_BuddyList
        .Checked = Not .Checked
        BuddyList.Visible = .Checked
    End With
End Sub

Private Sub mnu_View_Debug_Click()
    With mnu_View_Debug
        .Checked = Not .Checked
        DebugWin.Visible = .Checked
    End With

End Sub


Private Sub mnu_view_flatarea_Click()
    mnu_view_flatarea.Checked = True
    mnu_view_3darea.Checked = False
    Client.Appearance = 0
End Sub

Private Sub mnu_view_ForAWPos_Click()
    If intActive = -1 Then Exit Sub
    
    Dim strTemp As String, strPos As String, strCPos As String, strLst() As String
    strTemp = GetWindowTitle(intActive)
    
    PutINI winINI, "pos", "*" & strTemp, "-1,-1,-1,-1"
End Sub

Private Sub mnu_view_ForClientPos_Click()
    PutINI winINI, "pos", "!client", "-1,-1,-1,-1"
End Sub

Private Sub mnu_view_RemAWPos_Click()
    If intActive = -1 Then Exit Sub
    
    Dim strTemp As String, strPos As String
    strTemp = GetWindowTitle(intActive)
    
    With Client.ActiveForm
        strPos = .left & "," & _
                 .top & "," & _
                 .Width & "," & _
                 .Height
    End With
    
    PutINI winINI, "pos", "*" & strTemp, strPos
End Sub

Private Sub mnu_view_RemClientPos_Click()
    PutINI winINI, "pos", "!client", Client.left & "," & Client.top & "," & Client.Width & "," & Client.Height
End Sub

Private Sub mnu_view_ResClientPos_Click()
    SetClientPos
End Sub


Public Sub mnu_view_ResetAWPos_Click()
    If intActive = -1 Then Exit Sub
    
    Dim strTemp As String, strPos As String, strCPos As String, strLst() As String
    strTemp = GetWindowTitle(intActive)
    
    With Client.ActiveForm
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
    Client.ActiveForm.Move CInt(strLst(0)), CInt(strLst(1)), CInt(strLst(2)), CInt(strLst(3))

End Sub

Private Sub mnu_View_Status_Click()
    mnu_View_Status.Checked = Not mnu_View_Status.Checked
    Status.Visible = mnu_View_Status.Checked
End Sub

Private Sub mnu_viewTBTop_Click()
    mnu_view_TBTop.Checked = True
End Sub

Private Sub mnu_View_TBBot_Click()
    mnu_view_TBTop.Checked = False
    mnu_View_TBBot.Checked = True
    picTask.Align = 2
    If picToolMain.Align = 2 Then
        picToolMain.Align = 1
        picToolMain.Align = 2
    End If
End Sub

Private Sub mnu_View_TBrBot_Click()
    mnu_View_TBrTop.Checked = False
    mnu_View_TBrBot.Checked = True
    
    picToolMain.Align = 2
    If picTask.Align = 2 Then
        picTask.Align = 1
        picTask.Align = 2
    End If
End Sub

Private Sub mnu_View_TBrTop_Click()
    mnu_View_TBrTop.Checked = True
    mnu_View_TBrBot.Checked = False
    picToolMain.Align = 1
    If picTask.Align = 1 Then
        picTask.Align = 2
        picTask.Align = 1
    End If
End Sub


Private Sub mnu_view_TBTop_Click()
    mnu_view_TBTop.Checked = True
    mnu_View_TBBot.Checked = False
    picTask.Align = 1
    'If picToolMain.Align = 1 Then
    '    picToolMain.Align = 2
    '    picToolMain.Align = 1
    'End If
End Sub


Private Sub mnu_Window_Cascade_Click()
    Client.Arrange vbCascade
End Sub

Private Sub mnu_Window_TileH_Click()
    Client.Arrange vbTileHorizontal
End Sub


Private Sub picTask_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > minSize Then Exit Sub
    Dim which As Integer, wincnt As Integer, wID As Integer
    wincnt = WindowCount
    
    wID = minSize \ wincnt
    which = Int((X \ wID) + 0.5) + 1
    sngLastX = X
    
    If intActive = which Then
        intActive = 0
        HideWin which
'        DoEvents
        picTask.Tag = ""
        intHover = which
        DrawToolbar
        Exit Sub
    End If
    '* Which now contains which button was clicked
    intActive = which
    'DrawToolbar
    SetWinFocus which
    DrawToolbar
    SetWinFocus which
End Sub


Private Sub picTask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > minSize Then
        If intHover > 0 Then
            intHover = 0
            Client.DrawToolbar
            picTask.Tag = 0
            picTask.ToolTipText = ""
        End If
        Exit Sub
    End If
    
    Dim which As Integer, wincnt As Integer, wID As Integer
    wincnt = WindowCount
    
    wID = minSize \ wincnt
    which = Int((X \ wID) + 0.5) + 1
    'sngLastX = X
    intHover = which
    tmrTask.Enabled = True
    If Val(picTask.Tag) = intHover Then Exit Sub
    Client.DrawToolbar
    picTask.ToolTipText = strHoverCap
    
        
    'bReDraw = False
End Sub


Private Sub picToolMain_DblClick()
    If picToolMain.Tag = "" Then
        Load toolbardock
        toolbardock.Show
        SetParent picToolMain.hWnd, toolbardock.hWnd
        bDock = False
        picToolMain.Tag = ""
    Else
        SetParent picToolMain.hWnd, Client.hWnd
        Unload toolbardock
        bDock = True
        picToolMain.Tag = "dock"
    End If
    
    WriteINI "interface", "dock", CStr(bDock)
End Sub


Private Sub sock_Close()
    PutData Status.DataIn, strColor & "02Disconnected by SERVER from " & strServer
    sock.Close
    ClearModes
    IDENT.Close
    mnu_File_Connect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    mnu_File_Disconnect.Enabled = False
    Status.lblServer = "not connected"
    Status.Update
    BuddyList.lstNicks.Clear
    
    If bReconnect Then
        intCTries = 1
        PutData Status.DataIn, strColor & "02* Attempting to re-connect to server"
        Call mnu_File_Connect_Click
    End If

End Sub

Private Sub sock_Connect()
    '* Let's close all open windows
    Dim i As Integer
    For i = 1 To intChannels
        Channels(i).Tag = "NOPART"
        Unload Channels(i)
    Next i
    intChannels = 0
    
    For i = 1 To intQueries
        Unload Queries(i)
    Next i
    intQueries = 0
    mnu_File_Connect.Enabled = False
    mnu_File_Disconnect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    PutData Status.DataIn, strColor & "03Connected to " & strServer
    
    Dim strSendOut As String
    strSendOut = "PASS password" & vbCrLf & _
                 "NICK " & strMyNick & vbCrLf & _
                 "USER " & strMyNick & " local irc :" & strFullName
    SendData strSendOut
    
    Status.lblServer = sock.RemoteHost
    'strserver = sock.RemoteHost
    Status.Update
    
    '* Buddy List
    If mnu_View_BuddyList.Checked Then
        Dim strGet As String
        For i = 1 To BuddyList.lstSetup.ListCount
            strGet = strGet & BuddyList.lstSetup.List(i - 1) & " "
        Next i
        TimeOut 5
        If Trim(strGet) <> "" Then Client.SendData "ISON " & strGet

    End If
    
    
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String, AllParams As String
    Dim strData() As String, i As Integer
    
    lngPingReply = GetTickCount
    
    On Error Resume Next
    If sock.State <> 0 Then sock.GetData dat, vbString
    If DebugWin.Visible Then
        If Len(DebugWin.txtDataIn.Text) + Len(dat) > 60000 Then DebugWin.txtDataIn.Text = Right(DebugWin.txtDataIn.Text, 10000)
        On Error Resume Next
        If InStr(dat, "303 " & strMyNick & " :") Then Else DebugWin.txtDataIn = DebugWin.txtDataIn & "<< INCOMING DATA << " & vbCrLf
        DebugWin.txtDataIn.SelStart = Len(DebugWin.txtDataIn)
    End If
    
    strDataBuffer = strDataBuffer & dat
    
    If Right(strDataBuffer, 1) <> chr(13) And Right(strDataBuffer, 1) <> chr(10) Then
        Exit Sub
    Else
        dat = strDataBuffer
        strDataBuffer = ""
    End If
    
    strData = Split(dat, chr(10))
    
    For i = LBound(strData) To UBound(strData)
        interpret strData(i)
    Next i
    
End Sub


Private Sub sock_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData Status.DataIn, strColor & "04ERROR : " & description
   '* Cancel
    IDENT.Close
    sock.Close
    mnu_File_Connect.Caption = "&Connect"
    
    TimeOut 0.1

    If bRetry Then
        inttries = inttries + 1
        
        If intCTries >= intRetry Then
            intCTries = 1
            Dim s As String
            s = "ies"
            If intRetry = 1 Then s = "y"
            PutData Status.DataIn, strColor & "04* Failed to connect after " & intRetry & " tr" & s
        Else
            intCTries = intCTries + 1
            PutData Status.DataIn, strColor & "02* Attempting to connect to server again (attempt #" & intCTries & ")"
            Call mnu_File_Connect_Click
        End If
        
    End If
End Sub


Private Sub tmrTask_Timer()
    Dim pt As POINTAPI, lngRet As Long, hWnd As Long
    lngRet = GetCursorPos(pt)
    
    hWnd = WindowFromPoint(pt.X, pt.Y)
    If hWnd <> picTask.hWnd Then
        intHover = -1
        Client.DrawToolbar
        picTask.Tag = -1
        tmrTask.Enabled = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strNck As String, strSend As String
    Select Case Button.Index
        Case 1
            mnu_File_Connect_Click
        Case 2
            mnu_File_Disconnect_Click
        Case 3
            mnu_File_Options_Click
        Case 5
            Dim strInput As String
            strInput = InputX("Enter channel(s) to join, seperated each by a comma" & vbCrLf & "(EX: ""#irc"" or ""#irc, #mIRC"")." & vbCrLf & vbCrLf & "Also if you don't prefix # or &&, I will for you (EX: ""irc,help"")", "Join channels")
            If strInput = chr(8) Then Exit Sub
            strInput = GetAlias("", "JOIN " & strInput)
        Case 6
            Dim strMsg As String
            strNck = InputX("Enter the nickname of the user you would like to query", "Query user")
            If strNck = chr(8) Then Exit Sub
            strMsg = InputX("Enter the Message you would like to send to " & strNck & ".", "Query User Message")
            If strNck = chr(8) Then Exit Sub
            strSend = GetAlias("Status", "QUERY " & strNck & " " & strMsg)
            Client.SendData strSend
        Case 7  'dcc send
            strNck = InputX("Enter the nickname of the user you would like to DCC a file to", "DCC Send")
            If strNck = chr(8) Then Exit Sub
            strSend = GetAlias("", "DCCSEND " & strNck)
            Client.SendData strSend
    End Select
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            If Button.Value = tbrPressed Then
                ToggleMenu True
            Else
                ToggleMenu False
            End If
            
        Case 3
            PopupMenu mnu_File, , 572, 410
        Case 4
            PopupMenu mnu_view, , 972, 410
        Case 5
            PopupMenu mnu_window, , 1272, 410
        Case 6
            PopupMenu mnu_Help, , 1672, 410
    End Select
End Sub


Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            If Button.Value = tbrPressed Then
                ToggleMenu True
            Else
                ToggleMenu False
            End If
            
        Case 3
            PopupMenu mnu_File ', , 572, 410
        Case 4
            PopupMenu mnu_view ', , 972, 410
        Case 5
            PopupMenu mnu_window ', , 1272, 410
        Case 6
            PopupMenu mnu_Help ', , 1672, 410
    End Select

End Sub


