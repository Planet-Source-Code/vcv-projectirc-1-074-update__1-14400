VERSION 5.00
Begin VB.Form DebugWin 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debug"
   ClientHeight    =   2745
   ClientLeft      =   8700
   ClientTop       =   4335
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDataIn 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "IBMPC"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2775
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmDebugWin.frx":0000
      Top             =   0
      Width           =   8235
   End
End
Attribute VB_Name = "DebugWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCommandLine As Boolean
Public strBuff As String

Private Sub Form_Load()
    bCommandLine = True
    txtDataIn = "irc:\> "
    strBuffer = ""
End Sub

Private Sub Form_Resize()
    txtDataIn.Move txtDataIn.left, txtDataIn.top, Me.Width - 90, Me.Height - 330
End Sub


Private Sub txtDataIn_KeyPress(KeyAscii As Integer)

    If KeyAscii = 8 And bCommandLine Then
        If strBuff = "" Then
            KeyAscii = 0
        Else
            strBuff = left(strBuff, Len(strBuff) - 1)
            txtDataIn.SelStart = Len(txtDataIn) - 1
            txtDataIn.SelLength = 1
            txtDataIn.SelText = ""
            KeyAscii = 0
        End If
        Exit Sub
    End If
    
    If Not bCommandLine And KeyAscii <> 13 Then
        bCommandLine = True
        On Error Resume Next
        txtDataIn.SelStart = Len(txtDataIn)
        txtDataIn.SelText = "irc:\> " & chr(KeyAscii)
        strBuff = strBuff & chr(KeyAscii)
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        txtDataIn.SelStart = Len(txtDataIn)
        txtDataIn.SelText = vbCrLf
        If bCommandLine Then
            Client.SendData strBuff & vbCrLf
            bCommandLine = False
            txtDataIn.SelStart = Len(txtDataIn)
            txtDataIn.SelText = txtDataIn & vbCrLf
        Else
            bCommandLine = True
            txtDataIn.SelStart = Len(txtDataIn)
            txtDataIn.SelText = txtDataIn & "irc:\> "
        End If
        strBuff = ""
        
        Exit Sub
    End If
    
    txtDataIn.SelStart = Len(txtDataIn)
    txtDataIn.SelText = chr(KeyAscii)
    strBuff = strBuff & chr(KeyAscii)
End Sub


