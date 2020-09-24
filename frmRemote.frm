VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form RemoteEditor 
   Caption         =   "Remote Events Script Editor"
   ClientHeight    =   3705
   ClientLeft      =   3165
   ClientTop       =   2535
   ClientWidth     =   7365
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
   ScaleHeight     =   3705
   ScaleWidth      =   7365
   Begin RichTextLib.RichTextBox rtfScript 
      Height          =   3135
      Left            =   2520
      TabIndex        =   1
      Top             =   75
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRemote.frx":0000
   End
   Begin VB.ListBox lstEvents 
      Height          =   3105
      IntegralHeight  =   0   'False
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   2415
   End
End
Attribute VB_Name = "RemoteEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
