VERSION 5.00
Begin VB.Form ColorPicker 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Choose Color"
   ClientHeight    =   480
   ClientLeft      =   1830
   ClientTop       =   5820
   ClientWidth     =   2640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   2310
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   15
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   1980
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   14
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   1650
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   13
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   1320
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   12
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   990
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   11
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   660
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   10
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   330
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   9
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   8
      Top             =   240
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   2310
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   7
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   1980
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   6
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1650
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   5
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   1320
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   4
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   990
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   660
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   330
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.PictureBox picColor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   0
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If IsNumeric(chr(KeyAscii)) Then
        On Error Resume Next
        Client.ActiveForm.DataOut.SelText = chr(KeyAscii)
    Else
        On Error Resume Next
        Client.ActiveForm.DataOut.SelText = chr(KeyAscii)
        'Call Client.ActiveForm.DataOut_KeyPress(KeyAscii)
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 15
        picColor(i).BackColor = AnsiColor(i)
        picColor(i).Print i
    Next i
End Sub


Private Sub picColor_Click(Index As Integer)
    Client.ActiveForm.DataOut.SelText = Index
    Unload Me
End Sub

