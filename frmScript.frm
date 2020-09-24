VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form ExecScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Execute Script (BETA)"
   ClientHeight    =   3825
   ClientLeft      =   9015
   ClientTop       =   1185
   ClientWidth     =   5550
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   1305
      Top             =   3225
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
      UseSafeSubset   =   -1  'True
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "&Execute!"
      Height          =   360
      Left            =   4380
      TabIndex        =   1
      Top             =   3420
      Width           =   1080
   End
   Begin VB.TextBox txtSrc 
      Height          =   3330
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmScript.frx":0000
      Top             =   30
      Width           =   5445
   End
End
Attribute VB_Name = "ExecScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExec_Click()
  '  Dim i As Integer, strLine() As String
   ' strLine = Split(txtSrc.Text, vbCrLf)

    On Error Resume Next
    Script.ExecuteStatement txtSrc
End Sub

Private Sub Form_Load()
    Script.AddObject "Client", Client, True
End Sub


