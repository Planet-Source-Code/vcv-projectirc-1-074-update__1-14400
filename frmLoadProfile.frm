VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form LoadProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose User Profile"
   ClientHeight    =   3150
   ClientLeft      =   8880
   ClientTop       =   4230
   ClientWidth     =   3990
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
   ScaleHeight     =   3150
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin HoverButton.Button Button1 
      Height          =   2175
      Left            =   1470
      TabIndex        =   9
      Top             =   450
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   3836
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
      HilightColor    =   -2147483628
      ShadowColor     =   -2147483632
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "Button"
      CaptionDown     =   "Button"
      CaptionOver     =   "&Button"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   30
         ScaleHeight     =   2115
         ScaleWidth      =   2310
         TabIndex        =   10
         Top             =   30
         Width           =   2310
         Begin VB.ListBox lstUsers 
            Height          =   2175
            IntegralHeight  =   0   'False
            Left            =   -30
            TabIndex        =   11
            Top             =   -30
            Width           =   2385
         End
      End
   End
   Begin HoverButton.Button btnLoad 
      Height          =   375
      Left            =   2700
      TabIndex        =   8
      Top             =   2715
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
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
      HoverForeColor  =   8388608
      Caption         =   "Load Profile"
      CaptionDown     =   "&Load Profile"
      CaptionOver     =   "&Load Profile"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button btnCheckSkip 
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   2805
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
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
      HoverForeColor  =   0
      Caption         =   ""
      CaptionDown     =   "b"
      CaptionOver     =   ""
      ShowFocusRect   =   -1  'True
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   1
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button btnNew 
      Height          =   435
      Left            =   165
      TabIndex        =   4
      Top             =   450
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   767
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
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   8388608
      Caption         =   "Create New"
      CaptionDown     =   "Create New"
      CaptionOver     =   "&Create New"
      ShowFocusRect   =   -1  'True
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.Timer tmrDelete 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3105
      Top             =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   " &Skip this step on next load "
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   2790
      Width           =   2325
   End
   Begin HoverButton.Button btnRemove 
      Height          =   435
      Left            =   165
      TabIndex        =   5
      Top             =   1005
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   767
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
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   8388608
      Caption         =   "Remove Sel."
      CaptionDown     =   "Remove Sel."
      CaptionOver     =   "&Remove Sel."
      ShowFocusRect   =   -1  'True
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button btnDelete 
      Height          =   435
      Left            =   165
      TabIndex        =   6
      Top             =   2190
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   767
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
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   8388608
      Caption         =   "Cancel"
      CaptionDown     =   "Cancel"
      CaptionOver     =   "&Cancel"
      ShowFocusRect   =   -1  'True
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   930
      TabIndex        =   3
      Top             =   1890
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deleting Profile in ..."
      Height          =   495
      Left            =   150
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose user profile to use from below &:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   3330
   End
End
Attribute VB_Name = "LoadProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub LoadList()
    Dim strData As String, strUsers() As String
    
    On Error Resume Next
    
    '* if doesnt exist, create
    If FileExists(path & "Users.profile.list") = False Then
        Open path & "Users.profile.list" For Output As #1
            Print #1, ""
        Close #1
    End If
    
    '* open to read
    Open path & "Users.profile.list" For Binary As #1
        strData = String(LOF(1), 0)
        Get #1, 1, strData
    Close #1
    If Err Then
        strUserProfile = "NewUser"
        MsgBox "After trying to load the user profile list, an error has occured and the list was not loaded, and you cannot use user-profiles.  A profile under the name 'NewUser' was created until you can solve this problem." & vbCrLf & "Error #" & Err & " : " & Error, vbCritical
        Exit Sub
    End If
    
    strUsers = Split(strData, vbCrLf)
    Dim i As Integer
    For i = LBound(strUsers) To UBound(strUsers)
        If strUsers(i) <> "" Then lstUsers.AddItem strUsers(i)
    Next i
End Sub


Sub SaveList()
    Dim i As Integer, strLst As String
    For i = 0 To lstUsers.ListCount - 1
        strLst = strLst & lstUsers.List(i) & vbCrLf
    Next i
    
'    On Error Resume Next
    Open path & "Users.profile.list" For Output As #1
        Print #1, strLst
    Close #1
    
    If Err Then MsgBox "After trying to save the user-list, an error has occured, and the list was not able to be saved." & vbCrLf & "Error #" & Err & " : " & Error, vbCritical
    
End Sub


Private Sub btnDelete_Click()
        lstUsers.Enabled = True
        btnDelete.Visible = False
        lbl1.Visible = False
        lblCount.Visible = False
        lblCount.Caption = "3"
        tmrDelete.Enabled = False

End Sub


Private Sub btnLoad_Click()
        
    SaveList
    
    If lstUsers.ListCount = 0 Then
        MsgBox "There is no user profiles created, and to continue using this client, you must create one.  The only thing you need to do is click the 'Create New User' button and enter the name, and it does the rest.", vbExclamation
        Exit Sub
    End If
    
    If lstUsers.ListIndex = -1 Then
        MsgBox "You need to select a profile to load, please do this before continuing.", vbExclamation
        Exit Sub
    End If
    
    strUserProfile = lstUsers.List(lstUsers.ListIndex)
    
    Unload Me
    Client.Show

End Sub


Private Sub btnDel_Click()

End Sub

Private Sub btnCancel_Click()

End Sub

Private Sub btnNew_Click()
    Dim strUser As String
    strUser = InputX("Enter a name for the new user profile (can contain spaces): ", "New User Profile")
    
    If strUser = chr(8) Then Exit Sub
    
    If OnList(lstUsers, strUser) Then
        MsgBox "The name of the user profile you wish to create already exists.  2 user profiles cannot have the same name please try again with a different name.", vbInformation
        Exit Sub
    End If
    
    lstUsers.AddItem strUser
    Open path & strUser & "-settings.ini" For Output As #1
        Print ""
    Close #1

End Sub

Private Sub btnRemove_Click()
    If lstUsers.ListIndex = -1 Then
        MsgBox "You need to select a profile to delete, please do this before continuing.", vbExclamation
        Exit Sub
    End If
    
    Dim intRet As Integer
    intRet = MsgBox("Are you sure you would like to delete the selected profile? All settings will be permeanately lost.", vbYesNo Or vbQuestion)
    lstUsers.Enabled = False
    If intRet = vbYes Then
        btnDelete.Visible = True
        lbl1.Visible = True
        lblCount.Visible = True
        lblCount.Caption = 3
        tmrDelete.Enabled = True
    End If

End Sub


Private Sub cmdCancel_Click()
        cmdCancel.Visible = False
        lbl1.Visible = False
        lblCount.Visible = False
        lblCount.Caption = "3"
        tmrDelete.Enabled = False
        
End Sub

Private Sub cmdLoad_Click()
    
    SaveList
    
    If lstUsers.ListCount = 0 Then
        MsgBox "There is no user profiles created, and to continue using this client, you must create one.  The only thing you need to do is click the 'Create New User' button and enter the name, and it does the rest.", vbExclamation
        Exit Sub
    End If
    
    If lstUsers.ListIndex = -1 Then
        MsgBox "You need to select a profile to load, please do this before continuing.", vbExclamation
        Exit Sub
    End If
    
    strUserProfile = lstUsers.List(lstUsers.ListIndex)
    
    Unload Me
    Client.Show

End Sub


Private Sub cmdNewUser_Click()
    Dim strUser As String
    strUser = InputX("Enter a name for the new user profile (can contain spaces): ", "New User Profile")
    
    If strUser = chr(8) Then Exit Sub
    
    If OnList(lstUsers, strUser) Then
        MsgBox "The name of the user profile you wish to create already exists.  2 user profiles cannot have the same name please try again with a different name.", vbInformation
        Exit Sub
    End If
    
    lstUsers.AddItem strUser
    
End Sub


Private Sub Command1_Click()
    If lstUsers.ListIndex = -1 Then
        MsgBox "You need to select a profile to delete, please do this before continuing.", vbExclamation
        Exit Sub
    End If
    
    Dim intRet As Integer
    intRet = MsgBox("Are you sure you would like to delete the selected profile? All settings will be permeanately lost.", vbYesNo Or vbQuestion)
    If intRet = vbYes Then
        cmdCancel.Visible = True
        lbl1.Visible = True
        lblCount.Visible = True
        lblCount.Caption = 3
        tmrDelete.Enabled = True
    End If
End Sub

Private Sub btnRem_Click()

End Sub

Private Sub Check1_Click()
    btnCheckSkip.State = Check1.Value
End Sub

Private Sub Form_Load()
    Center Me
    'load for inputbox temporarely
    lngLeftColor = &HA27E66
    lngRightColor = &H8000000F
    lngForeColor = vbWhite
    
    path = App.path
    If Right(App.path, 1) <> "\" Then path = path & "\"
    DoEvents
    LoadList
    
    
    If lstUsers.ListCount >= 1 Then
        lstUsers.ListIndex = 0
    End If
End Sub


Private Sub lstUsers_DblClick()
    btnLoad_Click
End Sub

Private Sub lstUsers_GotFocus()
    Button1.State = Down
    
End Sub

Private Sub lstUsers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnLoad_Click
    End If
End Sub


Private Sub lstUsers_LostFocus()
    Button1.State = up
End Sub

Private Sub tmrDelete_Timer()
    lblCount.Caption = Int(lblCount.Caption) - 1
    
    If lblCount.Caption = "0" Then
        btnDelete.Visible = False
        lbl1.Visible = False
        lblCount.Visible = False
        tmrDelete.Enabled = False
        
        On Error Resume Next
        Dim strWhat As String
        strWhat = lstUsers.List(lstUsers.ListIndex)
        
        Kill path & strWhat & "-settings.ini"
        Kill path & strWhat & "-windows.ini"
        lstUsers.RemoveItem lstUsers.ListIndex
        lstUsers.Enabled = True
        SaveList
    End If
End Sub


