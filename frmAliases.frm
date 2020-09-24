VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AliasesEditor 
   Caption         =   "Script Aliases"
   ClientHeight    =   4305
   ClientLeft      =   7935
   ClientTop       =   1650
   ClientWidth     =   6435
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
   LockControls    =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   6435
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   315
      Left            =   5490
      TabIndex        =   5
      Top             =   3960
      Width           =   900
   End
   Begin MSComctlLib.ListView lvAliases 
      Height          =   3840
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   6773
      View            =   2
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Alias"
      Height          =   315
      Left            =   4425
      TabIndex        =   3
      Top             =   3960
      Width           =   1065
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   3840
      Left            =   2190
      TabIndex        =   2
      Top             =   60
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   6773
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAliases.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBMPC"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "–     &remove"
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   3945
      Width           =   1005
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+          &add"
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   3945
      Width           =   1005
   End
End
Attribute VB_Name = "AliasesEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bControl As Boolean
Public intSelected As Integer
Private Sub Command1_Click()
    If intSelected = -1 Then Exit Sub
    EditAlias lvAliases.ListItems.Item(intSelected).Text, lvAliases.ListItems.Item(intSelected).Text, rtfCode.Text
    
End Sub

Sub SaveAliases()
    Me.MousePointer = 11
    Dim i As Integer, strData As String
    
    For i = 1 To intAliasCnt
        strData = strData & _
                  Aliases(i).Name & Chr(8) & _
                  Aliases(i).Code
        If i = intAliasCnt Then
        Else
             strData = strData & Chr(0)
        End If
    Next i
    
    On Error Resume Next
    Open path & "aliases.data" For Output As #1
        Print #1, strData
    Close #1
    
    If Err Then
        MsgBox "While trying to save the aliases to a file, an error occrured." & _
                vbCrLf & "ERROR #" & Err & " : " & Error, vbCritical
    End If
    
    Me.MousePointer = 1

End Sub

Private Sub cmdAdd_Click()
    With lvAliases
        .ListItems.Add .ListItems.Count + 1, , "New_Alias", 0, 0
        .ListItems.Item(.ListItems.Count).Selected = True
        .SetFocus
        .StartLabelEdit
    End With
    AddAlias "New_Alias", "'* New_Alias : Add Description"
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    Dim ind As Integer
    ind = lvAliases.SelectedItem.Index
    
    RemoveAlias lvAliases.SelectedItem.Text '+ 1
    lvAliases.ListItems.Remove ind
    
End Sub


Private Sub cmdSave_Click()
    
    EditAlias lvAliases.SelectedItem, lvAliases.SelectedItem, rtfCode.Text
    
End Sub

Private Sub Form_Load()
    intSelected = -1
    
    Dim i As Integer
    For i = 1 To intAliasCnt
        lvAliases.ListItems.Add lvAliases.ListItems.Count + 1, , Aliases(i).Name, 0, 0
    Next i
    If intAliasCnt >= 1 Then
        lvAliases.ListItems.Item(1).Selected = True
        On Error Resume Next
        lvAliases.SetFocus
        rtfCode.Text = Aliases(1).Code
    End If
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 2000 Then Me.Height = 2000
    
    rtfCode.Width = Me.Width - 2355
    rtfCode.Height = Me.Height - 855
    lvAliases.Height = Me.Height - 855
    
    cmdAdd.top = lvAliases.Height + 115
    cmdRemove.top = cmdAdd.top
    cmdSave.top = cmdAdd.top
    cmdOk.top = cmdSave.top
    cmdOk.left = cmdSave.left + cmdSave.Width
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveAliases
End Sub

Private Sub lvAliases_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    '* Change alias name
    'MsgBox "!!" & lvAliases.SelectedItem & "~" & NewString
    EditAlias lvAliases.SelectedItem, NewString
    rtfCode.SetFocus
    rtfCode.SelStart = 0
    rtfCode.SelLength = Len(rtfCode.Text)
End Sub

Private Sub lvAliases_Click()
    On Error Resume Next
    rtfCode.Text = GetAliasCode(lvAliases.SelectedItem)
    
    If intSelected <> -1 Then
        On Error Resume Next
        lvAliases.ListItems(intSelected).bold = False
    End If
    
    intSelected = lvAliases.SelectedItem.Index
    lvAliases.SelectedItem.bold = True
'    rtfCode.SetFocus
    
    'Clipboard.SetText lvAliases.SelectedItem.Text
End Sub


Private Sub rtfCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = True   'control
    
    
End Sub

Private Sub rtfCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intAliasCnt > 0 Then cmdSave_Click
    ElseIf KeyAscii = 11 Then
        rtfCode.SelText = Chr(Color)
    ElseIf KeyAscii = 2 Then
        rtfCode.SelText = Chr(bold)
    ElseIf KeyAscii = 21 Then
        rtfCode.SelText = Chr(underline)
    ElseIf KeyAscii = 18 Then
        rtfCode.SelText = Chr(REVERSE)
    End If

End Sub


Private Sub rtfCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = False   'control
End Sub


