VERSION 5.00
Begin VB.Form LoadSettings 
   BorderStyle     =   0  'None
   ClientHeight    =   540
   ClientLeft      =   5055
   ClientTop       =   6300
   ClientWidth     =   3690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape2 
      BorderColor     =   &H00A27E66&
      BorderWidth     =   2
      Height          =   525
      Left            =   480
      Top             =   15
      Width           =   3210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A27E66&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   570
      Left            =   -15
      Top             =   -15
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Settings, Please wait ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   3300
   End
End
Attribute VB_Name = "LoadSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Center Me
End Sub


