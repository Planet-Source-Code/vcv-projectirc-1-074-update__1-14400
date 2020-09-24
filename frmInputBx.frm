VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form InputBx 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Box"
   ClientHeight    =   2295
   ClientLeft      =   8100
   ClientTop       =   5490
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInputBx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin HoverButton.Button Button2 
      Height          =   330
      Left            =   3600
      TabIndex        =   5
      Top             =   1920
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
      Caption         =   "Cancel"
      CaptionDown     =   "&Cancel"
      CaptionOver     =   "&Cancel"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Button1 
      Height          =   330
      Left            =   2730
      TabIndex        =   4
      Top             =   1920
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
      Caption         =   "Ok"
      CaptionDown     =   "&Ok"
      CaptionOver     =   "&Ok"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.PictureBox picInput 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1080
      ScaleHeight     =   270
      ScaleWidth      =   3420
      TabIndex        =   2
      Top             =   1560
      Width           =   3420
      Begin VB.TextBox txtInput 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   60
         TabIndex        =   3
         Text            =   "Default input"
         Top             =   30
         Width           =   3435
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00A27E66&
      BorderWidth     =   2
      Height          =   315
      Left            =   1035
      Top             =   1545
      Width           =   3495
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter input below:"
      Height          =   1410
      Left            =   1155
      TabIndex        =   1
      Top             =   75
      Width           =   3405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Input"
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
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   885
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A27E66&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1905
      Left            =   -15
      Top             =   -30
      Width           =   1110
   End
End
Attribute VB_Name = "InputBx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button1_Click()
    strIReturn = txtInput
    Unload Me

End Sub

Private Sub cmdCancel_Click()
    strIReturn = chr(8)
    Unload Me
End Sub

Private Sub cmdOK_Click()
    strIReturn = txtInput
    Unload Me
End Sub

Private Sub Button2_Click()
    strIReturn = chr(8)
    Unload Me

End Sub

Private Sub Form_Load()
    Center Me
    Me.Caption = strICaption
    lblText = strIText
    txtInput = strIDefault
    DoEvents
    SendMessage txtInput.hWnd, WM_SETFOCUS, 0&, 0
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput)
    
    Me.BackColor = lngRightColor
    Shape1.BackColor = lngLeftColor
    Shape2.BorderColor = lngLeftColor
    'SetButton Button1
    'SetButton Button2
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Button1_Click   'click ok
    End If
End Sub


