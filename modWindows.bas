Attribute VB_Name = "modWindows"
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const EM_UNDO = &HC7

'Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'Public Const WM_LBUTTONUP = &H202

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Global Const ICON_SIZE = 16
Global bSBL As Boolean

Private Const SWP_NOACTIVATE& = &H10

Private Const SWP_NOMOVE& = &H2

Private Const SWP_NOSIZE& = &H1

Private Const SWP_SHOWWINDOW& = &H40

Private Const HWND_BOTTOM& = 1
Private Const HWND_BROADCAST& = -1
Private Const HWND_DESKTOP& = 0
Private Const HWND_NOTOPMOST& = -2
Private Const HWND_TOP& = 0
Private Const HWND_TOPMOST& = -1

Sub Center(frm As Form)
    frm.Move (Screen.Width - frm.ScaleWidth) / 2, (Screen.Height - frm.ScaleHeight) / 2
End Sub

'
Public Sub FormOnTop(Form As Form)

        SetWindowPos Form.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE

End Sub
Sub DrawBGImage()
    If FileExists(strBGImage) = False Then Exit Sub
    
    Dim lngRet As Long
    On Error Resume Next
    Client.picBGImage.Picture = LoadPicture(strBGImage)
    'lngret = BitBlt(client
End Sub


Sub HideWin(intWhich As Integer)
    Dim i As Integer, cnt As Integer

    If intWhich = 1 Then
        If Status.Visible = False Then Status.Visible = True
        Status.Visible = False
        Exit Sub
    End If
    If intWhich = 2 Then
        If BuddyList.Visible = False Then BuddyList.Visible = True
        BuddyList.Visible = False
        Exit Sub
    End If
    
    'MsgBox intWhich
    cnt = 3
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then
            If cnt = intWhich Then
                Channels(i).Visible = False
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    For i = 1 To intQueries
        On Error Resume Next
        If Queries(i).strNick <> "" Then
            If cnt = intWhich Then
                Queries(i).Visible = False
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    For i = 1 To intDCCSends
        On Error Resume Next
        If DCCSends(i).Id <> "" Then
            If cnt = intWhich Then
                DCCSends(i).Visible = False
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    
    'WindowCount = cnt + 1 'add 1 for status window

End Sub

Public Function OnList(lstBox As ListBox, strItem As String) As Boolean
    Dim i As Integer
    For i = 0 To lstBox.ListCount - 1
        If LCase(lstBox.List(i)) = LCase(strItem) Then
            OnList = True
            Exit Function
        End If
    Next i
    OnList = False
End Function

Sub SetButton(Btn As Control)
'    Btn.BorderColor = lngLeftColor
'    Btn.BackColor = lngRightColor
'    Btn.HoverBackColor = lngRightColor
'    Btn.HilightColor = lngRightColor
'    Btn.ShadowColor = lngRightColor
End Sub

Public Function ShowAcceptDCC(strNick As String, strFile As String, lngSize As Long) As Boolean
    strDCCFile = strFile
    strDCCNick = strNick
    lngDCCSize = lngSize
    AcceptDCC.Show vbModal
    ShowAcceptDCC = bAcceptDCC
End Function


Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
    
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer


    With frmForm
        iLeft = .left / Screen.TwipsPerPixelX
        iTop = .top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
    


    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    Call SetWindowPos(frmForm.hWnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub
Function GetWindowIndex(strCaption As String)
    Dim i As Integer
    For i = 1 To WindowCount
        If GetWindowTitle(i) = strCaption Then
            GetWindowIndex = i
            Exit Function
        End If
    Next i
    GetWindowIndex = -1
End Function

Function GetWindowTitle(intWhich As Integer) As String
    Dim i As Integer, cnt As Integer
    If intWhich = 1 Then GetWindowTitle = "Status": Exit Function
    If intWhich = 2 Then GetWindowTitle = "Friend Tracker": Exit Function
    
    cnt = 2
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then cnt = cnt + 1
        If cnt = intWhich Then
            GetWindowTitle = Channels(i).strName
            Exit Function
        End If
    Next i
    
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then cnt = cnt + 1
        If cnt = intWhich Then
            On Error Resume Next
            GetWindowTitle = Queries(i).strNick
            Exit Function
        End If
    Next i
    
    For i = 1 To intDCCSends
        If DCCSends(i).Tag <> "" Then cnt = cnt + 1
        If cnt = intWhich Then
            GetWindowTitle = DCCSends(i).Caption
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
    For i = 1 To intDCCChats
        If DCCChats(i).Tag <> "" Then cnt = cnt + 1
        If cnt = intWhich Then
            GetWindowTitle = DCCChats(i).Caption
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
final:
    GetWindowTitle = ""
End Function
Sub SetWinFocus(intWhich As Integer)
    Dim i As Integer, cnt As Integer

    If intWhich = 1 Then
        If Status.Visible = False Then Status.Visible = True
        Status.SetFocus
        Status.Visible = True
        Exit Sub
    End If
    If intWhich = 2 Then
        If BuddyList.Visible = False Then BuddyList.Visible = True
        BuddyList.SetFocus
        BuddyList.Visible = True
        Exit Sub
    End If
    
    cnt = 2
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then
            cnt = cnt + 1
            If cnt = intWhich Then
                Channels(i).Visible = True
                Channels(i).Show
                Channels(i).WindowState = vbNormal
                SendMessage Channels(i).hWnd, WM_SETFOCUS, 0&, 0
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then
            cnt = cnt + 1
            If cnt = intWhich Then
                Queries(i).Visible = True
                Queries(i).WindowState = vbNormal
                Queries(i).DataOut.SelStart = 0
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To intDCCSends
        If DCCSends(i).Tag <> "" Then
            cnt = cnt + 1
            If cnt = intWhich Then
                DCCSends(i).Visible = True
                DCCSends(i).SetFocus
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To intDCCChats
        If DCCChats(i).Tag <> "" Then
            cnt = cnt + 1
            If cnt = intWhich Then
                DCCChats(i).Visible = True
                DCCChats(i).SetFocus
                Exit Sub
            End If
        End If
    Next i
    
End Sub

Function TaskCenter(intActual As Integer, strText As String) As Integer
    Dim intRet As Integer
    intRet = (((intActual - ICON_SIZE) - Client.picTask.TextWidth(strText)) / 2) + (ICON_SIZE / 2)
    If Right(strText, 1) = "." Then intRet = intRet + 4
    intRet = intRet + 4
    TaskCenter = intRet
End Function


Function TaskText(intWidth As Integer, strText As String) As String
    Dim lastWidth As Integer, i As Integer, strBuf As String
    Dim inttemp As Integer
    
    For i = 1 To Len(strText)
        strBuf = left(strText, i) & "..."
        inttemp = Client.picTask.TextWidth(strBuf) ' + 2 + ICON_SIZE
        
        If inttemp >= intWidth - 2 - 32 Then
            TaskText = left(strText, i - 1) & "..."
            Exit Function
        End If
    Next i
    TaskText = strText
        
End Function

Function WindowCount() As Integer
    Dim cnt As Integer, i As Integer
    
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then cnt = cnt + 1
    Next i
    
    For i = 1 To intQueries
        On Error Resume Next
        If Queries(i).strNick <> "" Then cnt = cnt + 1
    Next i
    
    For i = 1 To intDCCSends
        On Error Resume Next
        If DCCSends(i).Tag <> "" Then cnt = cnt + 1
    Next i
    
    For i = 1 To intDCCChats
        On Error Resume Next
        If DCCChats(i).Tag <> "" Then cnt = cnt + 1
    Next i
    
    WindowCount = cnt + 2 'add 2 for status window and buddy list
    
    '* Add DCC and stuff here

End Function


Function WindowNewBuffer(intWhich As Integer) As String
    Dim i As Integer, cnt As Integer
    If intWhich = 1 Then
        If Status.newBuffer = True Then
            WindowNewBuffer = True
        Else
            WindowNewBuffer = False
        End If
        Exit Function
    End If
    If intWhich = 1 Then
        If BuddyList.newBuffer = True Then
            WindowNewBuffer = True
        Else
            WindowNewBuffer = False
        End If
        Exit Function
    End If
    
    cnt = 3
    For i = 1 To intChannels
        If cnt = intWhich Then
            If Channels(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
            Else
                WindowNewBuffer = False
                Exit Function
                DoEvents
            End If
        End If
        cnt = cnt + 1
    Next i
    
    
    For i = 1 To intQueries
        If cnt = intWhich Then
            On Error Resume Next
            If Queries(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
            Else
                WindowNewBuffer = False
                Exit Function
            End If
        End If
        cnt = cnt + 1
    Next i
    
    GoTo final
    
    For i = 1 To intDCCChats
        If cnt = intWhich Then
'            If dccchats(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
'            Else
                WindowNewBuffer = False
                Exit Function
'            End If
        End If
        cnt = cnt + 1
    Next i
    

final:
    WindowNewBuffer = False ' ""

End Function


