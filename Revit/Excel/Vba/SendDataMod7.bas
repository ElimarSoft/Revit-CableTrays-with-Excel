Attribute VB_Name = "SendDataMod7"
Option Explicit
    
'https://www.vbforums.com/showthread.php?73053-How-can-i-send-an-quot-CTRL-quot-i-command-using-PostMessage-(or-SendMessage)
'https://learn.microsoft.com/es-es/windows/win32/api/winuser/nf-winuser-keybd_event
'https://learn.microsoft.com/es-es/windows/win32/api/winuser/nf-winuser-sendinput
'https://learn.microsoft.com/es-es/windows/win32/api/winuser/nf-winuser-getasynckeystate

Public Declare PtrSafe Function apiPostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare PtrSafe Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare PtrSafe Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function apiFindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare PtrSafe Function apiGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal szWindowText As String, ByVal lLength As Long) As Long
Public Declare PtrSafe Sub apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function apiSetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal Code As Long, ByVal MapType As Long) As Long

Const WM_INPUT = 255
Const WM_KEYDOWN = 256
Const WM_KEYFIRST = 256
Const WM_KEYUP = 257
Const WM_CHAR = 258
Const WM_DEADCHAR = 259
Const WM_SYSKEYDOWN = 260
Const WM_SYSKEYUP = 261
Const WM_SYSCHAR = 262
Const WM_SYSDEADCHAR = 263

Const VK_TAB = &H9
Const VK_RETURN = &HD
Const VK_SHIFT = &H10
Const VK_CONTROL = &H11
Const VK_MENU = &H12 'Alt Key
Const VK_F1 = &H70
Const VK_F2 = &H71
Const VK_F3 = &H72
Const VK_F4 = &H73
Const VK_F5 = &H74
Const VK_F6 = &H75
Const VK_F7 = &H76
Const VK_F8 = &H77
Const VK_F9 = &H78
Const VK_F10 = &H79
Const VK_F11 = &H7A
Const VK_F12 = &H7B

Public Sub ExecMacro2(Data As String)

    'Dim Data As String
    'Data = "GetCadData"
    Dim n As Integer
    Dim TimeOut2 As Integer: TimeOut2 = 50
    
    Dim h1 As Long
    h1 = FindRevit
    If h1 = 0 Then Exit Sub
    
    Call SendText(h1, "md") 'Enter Modify Mode
   
    Delay (2)
    DoEvents
        
'    Call apiPostMessage(h1, WM_SYSKEYDOWN, VK_MENU, &H20000001)
'    Call apiPostMessage(h1, WM_SYSKEYUP, VK_MENU, &HC0000001)
'    Delay (2)
'    Call SendText(h1, "GVP")
    
    Call SendSysControlKeys(h1, VK_F10, 1)
    Delay (2)
    Call SendText(h1, "G")
    Delay (2)
    Call SendText(h1, "VP")
    
    Do
        Delay (1)
        h1 = apiFindWindow("Chrome_WidgetWin_1", "Dynamo Player")
        If h1 > 0 Then Exit Do
        TimeOut2 = TimeOut2 - 1
        If TimeOut2 <= 0 Then Exit Do
    Loop
    
    If h1 > 0 Then
        Delay (5)
        Call SendControlKeys(h1, VK_TAB, 4)
        Delay (5)
        Call SendText(h1, Data)
        Call SendControlKeys(h1, VK_TAB, 6)
        Call SendControlKeys(h1, VK_RETURN, 1)
        Call SendControlKeys(h1, VK_TAB, 6)
        Call SendAltKey(h1, VK_F4)
    End If

End Sub
Private Sub SendSysControlKeys(h1 As Long, VK As Integer, count As Integer)
    Dim n As Integer
    For n = 1 To count
        Call apiPostMessage(h1, WM_SYSKEYDOWN, VK, &H20000001)
        Call apiPostMessage(h1, WM_SYSKEYUP, VK, &HC0000001)
    Next n
    DoEvents
End Sub
Private Sub SendControlKeys(h1 As Long, VK As Integer, count As Integer)
    Dim n As Integer
    For n = 1 To count
        Call apiPostMessage(h1, WM_KEYDOWN, VK, &H20000001)
        Call apiPostMessage(h1, WM_KEYUP, VK, &HC0000001)
    Next n
    DoEvents
End Sub
Private Sub SendText(h1 As Long, Data As String)
    Dim n As Integer
    For n = 1 To Len(Data)
        Call apiPostMessage(h1, WM_CHAR, Asc(Mid(Data, n, 1)), &H1)
    Next n
    DoEvents
End Sub
Private Sub SendAltChar(h1 As Long, VK As Integer)
    
    Call apiPostMessage(h1, WM_SYSKEYDOWN, VK_MENU, &H20000001)
    Call apiPostMessage(h1, WM_SYSCHAR, VK, &H1)
    DoEvents
    Call apiPostMessage(h1, WM_KEYUP, VK_MENU, &HC0000001)
    DoEvents

End Sub

Private Sub SendAltKey(h1 As Long, VK As Integer)
    
    Call apiPostMessage(h1, WM_SYSKEYDOWN, VK_MENU, &H20000001)
    DoEvents
    Delay (2)
    'Call apiPostMessage(h1, WM_SYSKEYDOWN, VK, &H203E0001)
    Call apiPostMessage(h1, WM_SYSKEYDOWN, VK, &H20000001)
    DoEvents
    Call apiPostMessage(h1, WM_KEYUP, VK_MENU, &HC0000001)
    DoEvents

End Sub


Private Function FindRevit() As Long
    Dim h1 As Long
    Const WindowName As String = "Autodesk Revit"
    Const ScreenName As String = "3D View"
    Const WinClass As String = "AfxFrameOrView140u"
    h1 = GetDesktopWindow()
    h1 = FindWindowByName(h1, WindowName)
    h1 = FindWindowByName(h1, ScreenName)
    h1 = FindWindowEx(h1, 0, WinClass, vbNullString)
    FindRevit = h1
End Function

Private Function FindWindowByName(h1 As Long, WindowName As String) As Long
    
    Dim h2 As Long
    Dim text As String * 255
    
    Do
        h2 = apiFindWindowEx(h1, h2, vbNullString, vbNullString)
        Call apiGetWindowText(h2, text, 255)
        If (Left(text, Len(WindowName)) = WindowName) Then
            Exit Do
        End If
        If h2 = 0 Then Exit Do
    Loop
    FindWindowByName = h2

End Function
Private Sub Delay(n As Integer)
    apiSleep (CInt(Range("TimeOut").Value) * n)
End Sub

