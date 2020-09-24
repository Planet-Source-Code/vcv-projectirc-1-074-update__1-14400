VERSION 5.00
Begin VB.Form tooltip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   975
   ClientTop       =   2655
   ClientWidth     =   4335
   FillColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   30
      TabIndex        =   3
      Top             =   15
      Width           =   4260
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Information regarding the current command goes right about here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   435
      Left            =   390
      TabIndex        =   2
      Top             =   285
      Width           =   3900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   270
      Width           =   165
   End
   Begin VB.Label lblCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/command <param1> <param2>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   210
      Left            =   300
      TabIndex        =   0
      Top             =   30
      Width           =   3990
   End
   Begin VB.Shape shp 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   765
      Left            =   30
      Top             =   30
      Width           =   225
   End
End
Attribute VB_Name = "tooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Me.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    'SendMessage Client.ActiveForm.DataOut.hwnd, WM_SETFOCUS, 0, 0
    'MsgBox "hehe"
    
    If trans.fIsWin2000() Then
        'If trans.fIsWin2000() Then trans.fSetTranslucency tooltip.hWnd, 255
    End If
End Sub


Private Sub Form_Resize()
    
    Me.Cls
On Error Resume Next
Line (0, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), &H808080, B
    Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), vbBlack
    Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), vbBlack
    shp.Height = Me.ScaleHeight - 3
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Visible = False
End Sub


