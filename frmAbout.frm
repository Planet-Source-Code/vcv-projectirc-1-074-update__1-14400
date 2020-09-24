VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3555
   ClientLeft      =   5115
   ClientTop       =   2385
   ClientWidth     =   5295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2100
      Left            =   630
      ScaleHeight     =   2100
      ScaleWidth      =   4305
      TabIndex        =   4
      Top             =   825
      Width           =   4305
      Begin RichTextLib.RichTextBox rtfAbout 
         Height          =   2205
         Left            =   45
         TabIndex        =   5
         Top             =   30
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   3889
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         FileName        =   "C:\vb\irc\about.rtf"
         TextRTF         =   $"frmAbout.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin HoverButton.Button btnClose 
      Height          =   390
      Left            =   4050
      TabIndex        =   6
      Top             =   3015
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   688
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
      Caption         =   "Close"
      CaptionDown     =   "&Close"
      CaptionOver     =   "&Close"
      ShowFocusRect   =   -1  'True
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.Label lblShell 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.projectIRC.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A27E66&
      Height          =   195
      Left            =   675
      MouseIcon       =   "frmAbout.frx":0B9D
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3105
      Width           =   2025
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00A27E66&
      BorderWidth     =   2
      Height          =   2145
      Left            =   615
      Top             =   810
      Width           =   4350
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  'Transparent
      Caption         =   "1.xxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   3765
      TabIndex        =   1
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0EA7
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   780
      TabIndex        =   2
      Top             =   1050
      Width           =   4005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "projectIRC "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A27E66&
      Height          =   675
      Left            =   630
      TabIndex        =   0
      Top             =   15
      Width           =   3300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   810
      Left            =   -615
      Top             =   -15
      Width           =   5625
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Button1_Click()

End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    
    lblVer = App.Major & "." & App.Minor & App.Revision
   
            
                                                                                                                                        Label3 = "©2000 Matt C, sappy@adelphia.net                              Do not use full source code w/o permission or take credit for anything within this program.                           Written in Visual Basic 6.0 Enterprise Edition."
                                                                                                                                        lblShell = "http://www.projectIRC.com"
End Sub

Private Sub lblShell_Click()
    ShellExecute 0, "open", "http://www.projectirc.com", "", "", 0
End Sub


