VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form BuddyList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Friend Tracker"
   ClientHeight    =   3630
   ClientLeft      =   11670
   ClientTop       =   2295
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBuddyList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   Begin HoverButton.Button cmdRefresh 
      Height          =   330
      Left            =   1830
      TabIndex        =   4
      Top             =   30
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   582
      BackColor       =   -2147483633
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   4210752
      Caption         =   "Refresh"
      CaptionDown     =   "&Refresh"
      CaptionOver     =   "&Refresh"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   105
      ScaleHeight     =   2745
      ScaleWidth      =   2550
      TabIndex        =   1
      Top             =   405
      Width           =   2550
      Begin VB.Timer tmrRefresh 
         Interval        =   10000
         Left            =   615
         Top             =   840
      End
      Begin VB.ListBox lstSetup 
         Height          =   2820
         IntegralHeight  =   0   'False
         Left            =   -30
         TabIndex        =   3
         Top             =   -30
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.ListBox lstNicks 
         Height          =   2805
         IntegralHeight  =   0   'False
         Left            =   -30
         TabIndex        =   2
         Top             =   -30
         Width           =   2610
      End
   End
   Begin HoverButton.Button cmdAdd 
      Height          =   285
      Left            =   2055
      TabIndex        =   5
      Top             =   3225
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      BackColor       =   -2147483633
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   4210752
      Caption         =   "+"
      CaptionDown     =   "+"
      CaptionOver     =   "+"
      ShowFocusRect   =   0   'False
      Sink            =   0   'False
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button cmdRem 
      Height          =   285
      Left            =   2370
      TabIndex        =   6
      Top             =   3225
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      BackColor       =   -2147483633
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   4210752
      Caption         =   "-"
      CaptionDown     =   "-"
      CaptionOver     =   "-"
      ShowFocusRect   =   0   'False
      Sink            =   0   'False
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2790
      Left            =   90
      Top             =   390
      Width           =   2595
   End
   Begin VB.Label lblWhich 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Online"
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
      Height          =   195
      Left            =   255
      TabIndex        =   0
      Top             =   75
      Width           =   990
   End
   Begin VB.Shape shpLeftC 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3660
      Left            =   0
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "BuddyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intMins As Integer
Public newBuffer As Boolean
Public Sub Reload()
    Call cmdRefresh_Click
End Sub

Sub AddUsers(strData As String)
    Dim names() As String, i As Integer, lngFind As Long, strData2 As String
    names = Split(strData, " ")
    
    For i = 0 To lstNicks.ListCount - 1
        If InStr(strData, lstNicks.List(i)) Then
        Else
            EchoActive lstNicks.List(i) & " has logged " & strBold & "off" & strBold & "line.", 7
        End If
    Next i
    
    For i = 0 To UBound(names)
        If names(i) <> "" Then
            'lb_findstring=&H18F
            lngFind = SendMessage(lstNicks.hWnd, &H18F, 0, ByVal names(i))
            Debug.Print names(i) & "~" & CStr(lngFind)
            If lngFind = -1 Then
                EchoActive names(i) & " has logged online.", 7
                lstNicks.AddItem names(i)
            End If
        End If
    Next i
    
    lstNicks.Clear
    
    For i = 0 To UBound(names)
        If names(i) <> "" Then lstNicks.AddItem names(i)
    Next i

    
End Sub


Sub LoadList()
    Dim names() As String, strData As String
    Dim i As Integer
    
    lstSetup.Clear
    
    On Error Resume Next
    Open path & "friends.trk" For Binary As #1
        strData = String(LOF(1), 0)
        Get #1, 1, strData
        DoEvents
    Close #1
    DoEvents
    
    names = Split(strData, vbCrLf)
    
    For i = 0 To UBound(names)
        If Trim(names(i)) <> "" Then lstSetup.AddItem names(i)
        DoEvents
    Next i
    
    If Err Then MsgBox "Error occured while loading Buddy List : " & Error
End Sub

Sub SaveList()
    Dim i As Integer, strFinal As String
    For i = 1 To lstSetup.ListCount
        strFinal = strFinal & lstSetup.List(i - 1) & vbCrLf
    Next i
    strFinal = LeftR(strFinal, 2)

    On Error Resume Next
    Open path & "Friends.trk" For Output As #1
        Print #1, strFinal
    Close #1
    If Err Then MsgBox "Error occured while saving Friend Tracker List : " & Error
End Sub

Private Sub cmdAdd_Click()
    Dim strName As String
    strName = InputBox("Enter name to add", "Friend Tracker", "")
    If strName = "" Then Exit Sub
    
    strName = Replace(strName, " ", "")
    Dim i As Integer
    For i = 0 To lstSetup.ListCount
        If LCase(lstSetup.List(i)) = LCase(strName) Then Exit Sub
    Next i
    
    lstSetup.AddItem Trim(strName)
    Call cmdRefresh_Click
    SaveList
End Sub


Private Sub cmdRefresh_Click()
    If Client.sock.State = 0 Then Exit Sub
    Dim i As Integer, strGet As String
    For i = 1 To lstSetup.ListCount
        strGet = strGet & lstSetup.List(i - 1) & " "
    Next i
    If Trim(strGet) <> "" Then Client.SendData "ISON " & strGet
End Sub

Private Sub cmdRem_Click()
    Dim ind As Integer
    ind = lstSetup.ListIndex
    If ind = -1 Then Exit Sub
    lstSetup.RemoveItem ind
    SaveList
End Sub

Private Sub Command1_Click()
    cmdAdd.Visible = False
    cmdRem.Visible = False
    lstSetup.Visible = False
    lblWhich.Caption = "Online"
End Sub


Private Sub Command2_Click()
    cmdAdd.Visible = True
    cmdRem.Visible = True
    lstSetup.Visible = True
    lblWhich.Caption = "Setup"
End Sub


Private Sub Form_Activate()
    If bSBL = True Then
        DoEvents
        If BuddyList.Visible = True Then BuddyList.Move Client.Width - BuddyList.Width - 180, 0
    End If
    Dim i As Integer
        Client.mnu_View_BuddyList.Checked = True
    i = GetWindowIndex("Friend Tracker")
    SetWinFocus i
    Client.intActive = i
    Client.intHover = -1
    newBuffer = False

    Client.DrawToolbar
End Sub

Private Sub Form_Load()
    
    LoadList
    bSBL = True
    Client.mnu_View_BuddyList.Checked = True
    cmdRefresh_Click
    lstNicks.BackColor = lngBackColor
    lstNicks.ForeColor = lngForeColor
    lstSetup.BackColor = lngBackColor
    lstSetup.ForeColor = lngForeColor
    
    
    '* Set the colors straight!!
    Me.BackColor = lngRightColor
    
    shpLeftC.BackColor = lngLeftColor
    shpBorder.BorderColor = lngLeftColor
    cmdRefresh.BorderColor = lngLeftColor
    cmdAdd.BorderColor = lngLeftColor
    cmdRem.BorderColor = lngLeftColor
    SetButton cmdAdd
    SetButton cmdRem
    SetButton cmdRefresh
    
    Me.Visible = True
    DoEvents
    Client.mnu_view_ResetAWPos_Click
    'BuddyList.lstNicks.Font.Name = strFontName
    'BuddyList.lstSetup.Font.Name = strFontName
    'lstNicks.Font.Size = intFontSize
    'lstSetup.Font.Size = intFontSize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Client.mnu_View_BuddyList.Checked = False
End Sub

Private Sub lstNicks_DblClick()
    NewQuery lstNicks.List(lstNicks.ListIndex), ""
End Sub

Private Sub tmrRefresh_Timer()
    Call cmdRefresh_Click
End Sub


