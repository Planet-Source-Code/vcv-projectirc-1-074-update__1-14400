VERSION 5.00
Begin VB.Form toolbardock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ToolBar"
   ClientHeight    =   435
   ClientLeft      =   2520
   ClientTop       =   1845
   ClientWidth     =   4440
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
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "toolbardock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    If Me.Tag = "" Then
        SetParent Client.picToolMain.hWnd, Client.hWnd
        Unload toolbardock
        bDock = True
        Client.picToolMain.Tag = "dock"
        WriteINI "interface", "dock", CStr(bDock)
    End If
    Unload Me
End Sub


