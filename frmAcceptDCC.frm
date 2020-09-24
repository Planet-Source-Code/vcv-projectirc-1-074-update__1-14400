VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form AcceptDCC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accept DCC"
   ClientHeight    =   2640
   ClientLeft      =   6540
   ClientTop       =   3375
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcceptDCC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAcceptDCC.frx":27A2
   ScaleHeight     =   2640
   ScaleWidth      =   4935
   Begin HoverButton.Button cmdBrowse 
      Height          =   300
      Left            =   4485
      TabIndex        =   13
      Top             =   2250
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   529
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
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
      BorderColor     =   -2147483632
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "..."
      CaptionDown     =   "..."
      CaptionOver     =   "..."
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button cmdAccept 
      Height          =   375
      Left            =   3705
      TabIndex        =   11
      Top             =   90
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      BorderColor     =   -2147483632
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "Accept DCC"
      CaptionDown     =   "Accept DCC"
      CaptionOver     =   "&Accept DCC"
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
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4290
      Picture         =   "frmAcceptDCC.frx":306C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   1050
      Width           =   480
   End
   Begin VB.TextBox txtFullPath 
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Text            =   "C:\projectIRC\<filename>"
      Top             =   2250
      Width           =   4305
   End
   Begin HoverButton.Button cmdDecline 
      Height          =   375
      Left            =   3705
      TabIndex        =   12
      Top             =   495
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
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
      BorderColor     =   -2147483632
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "Decline DCC"
      CaptionDown     =   "Decline DCC"
      CaptionOver     =   "&Decline DCC"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      Caption         =   "<file size>"
      Height          =   225
      Left            =   600
      TabIndex        =   7
      Top             =   1935
      Width           =   4215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Size &:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label lblFile 
      Alignment       =   1  'Right Justify
      Caption         =   "<file name>"
      Height          =   225
      Left            =   600
      TabIndex        =   5
      Top             =   1710
      Width           =   4215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "File &:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Would you like to Accept this DCC Transfer?"
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   1380
      Width           =   3165
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAcceptDCC.frx":580E
      Height          =   1020
      Left            =   90
      TabIndex        =   2
      Top             =   270
      Width           =   3465
   End
   Begin VB.Label lblNick 
      AutoSize        =   -1  'True
      Caption         =   "Unknown Nickname"
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
      Left            =   810
      TabIndex        =   1
      Top             =   75
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "The user, "
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   735
   End
   Begin VB.Label Label6 
      Height          =   1230
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3360
   End
End
Attribute VB_Name = "AcceptDCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccept_Click()
    
    If FileExists(txtFullPath) Then
        Dim intChoice As Integer, strSuggest
        If Val(lblSize) >= CLng(FileLen(txtFullPath)) Then
            strSuggest = "It appears that you already have the file, or the files conflict.  It is suggested that you click CANCEL and rename the file."
        Else
            strSuggest = "It appears that the file to be sent is smaller than the file on your Computer, and is suggest that you resume the file by click YES, if the files do not conflict."
        End If
        
        intChoice = MsgBox("You have chosen to accept the DCC Transfer, but the directory you chose to save the file contains a file of the same name already." & vbCrLf & "The size of the current file is " & Format(CStr(FileLen(txtFullPath)), "###,###,###,###") & " bytes, while the file you choose to save is " & _
                    Format(lblSize, "###,###,###,###") & " bytes." & vbCrLf & strSuggest & vbCrLf & "Click YES to Resume this file, NO to Overwrite it, and CANCEL to rename the file.", vbCritical Or vbYesNoCancel)
        If intChoice = vbYes Then
            lngDCCStart = FileLen(txtFullPath)
            If Val(lblSize) >= CLng(FileLen(txtFullPath)) Then
                lngDCCStart = -1
            End If
        ElseIf intChoice = vbNo Then
            lngDCCStart = -1
        Else
            Exit Sub
        End If
    Else
        lngDCCStart = -1
    End If
    bAcceptDCC = True
    strDCCFile = txtFullPath
    Unload Me
End Sub


Private Sub cmdBrowse_Click()
    Dim strFile As SelectedFile
    strFile = ShowSave(Me.hWnd)
    If strFile.bCanceled Then Exit Sub
    txtFullPath = strFile.sLastDirectory & strFile.sFiles(1)
End Sub

Private Sub cmdDecline_Click()
    bAcceptDCC = False
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim strFile As SelectedFile
    strFile = ShowSave(Me.hWnd)
    If strFile.bCanceled Then Exit Sub
    txtFullPath = strFile.sLastDirectory & strFile.sFiles(1)
End Sub

Private Sub Form_Load()
    lblNick = strDCCNick
    lblFile = strDCCFile
    lblSize = Format(CStr(lngDCCSize), "###,###,###,### bytes")
    txtFullPath = path & lblFile
End Sub


