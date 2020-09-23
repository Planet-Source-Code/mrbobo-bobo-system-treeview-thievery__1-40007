Attribute VB_Name = "ModBFF"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - October 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au
Option Explicit
Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const WM_MOVE = &H3
Private Const GWL_WNDPROC = (-4)
Private lpPrevWndProc As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_NEXT = 2
Private Const GW_CHILD = 5
Private Const WM_CLOSE = &H10
Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
Public m_CurrentDirectory As String
Dim DialogWindow As Long
Dim SysTreeWindow As Long
Dim CancelbuttonWindow As Long
Dim DialogContainer As Object

'Tandard BrowseForFolder dialog
Private Sub BrowseForFolder(StartDir As String)
    Dim lpIDList As Long
    Dim szTitle As String
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    m_CurrentDirectory = StartDir & vbNullChar
    With tBrowseInfo
        .hwndOwner = GetDesktopWindow
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_STATUSTEXT
        'We need to process messages
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
End Sub


Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    Dim lpIDList As Long
    Dim Ret As Long
    Dim sBuffer As String
    Dim hwnda As Long, ClWind As String * 14, ClCaption As String * 100
    On Error Resume Next
    DialogWindow = hwnd 'Handle of BrowseForFolder dialog
    Select Case uMsg
        Case BFFM_INITIALIZED
            'Move the whole  BrowseForFolder dialog off screen
            Call MoveWindow(DialogWindow, -Screen.Width, 0, 480, 480, True)
            'Set it's initial path
            Call SendMessage(hwnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
            'Enumerate cild windows
            hwnda = GetWindow(hwnd, GW_CHILD)
            Do While hwnda <> 0
                GetClassName hwnda, ClWind, 14
                'Found a button
                If Left(ClWind, 6) = "Button" Then
                    GetWindowText hwnda, ClCaption, 100
                    'If it's the Cancel button, remember it's
                    'handle so we can press it later
                    If UCase(Left(ClCaption, 6)) = "CANCEL" Then
                        CancelbuttonWindow = hwnda
                    End If
                End If
                'Here's what we're really after - it's Treeview!
                If Left(ClWind, 13) = "SysTreeView32" Then
                    SysTreeWindow = hwnda
                End If
                hwnda = GetWindow(hwnda, GW_NEXT)
            Loop
            'Steal the Treeview for our own use
            GrabTV DialogContainer
        Case BFFM_SELCHANGED
            'Path has changed - better tell our form
            sBuffer = Space(MAX_PATH)
            Ret = SHGetPathFromIDList(lp, sBuffer)
            m_CurrentDirectory = sBuffer
            Form1.PathChange
    End Select
    BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
    GetAddressofFunction = add
End Function
Private Sub GrabTV(mNewOwner As Object)
    'Thievery in progress
    Dim R As RECT
    'It's mine now!
    SetParent SysTreeWindow, mNewOwner.hwnd
    'Put it where we want it
    GetWindowRect mNewOwner.hwnd, R
    SizeTV 0, 0, mNewOwner.ScaleWidth, mNewOwner.ScaleHeight
    'Temporary hook to catch the move event
    DialogHook
End Sub
Public Sub CloseUp()
    'Send the Treeview back to the BrowseForFolder dialog
    SetParent SysTreeWindow, DialogWindow
    'Close the dialog
    SendMessage DialogWindow, WM_CLOSE, 1, 0
    'Just to be sure...
    DestroyWindow DialogWindow
End Sub
Private Sub TaskbarHide()
    'Hide the BrowseForFolder dialog from the Taskbar
    ShowWindow DialogWindow, 0
    'Done with hooking
    DialogUnhook
End Sub
Public Sub main()
    'Project startup routine required so that
    'our container is fully opened before we
    'use the Setparent API
    Form1.Show 'load up
    Set DialogContainer = Form1.PicBrowse 'container for the Treeview
    BrowseForFolder "c:\" 'Spawn the dialog
End Sub

'This hook is very temporary and is self-cancelling
'It is needed because we need to hide the BrowseForFolder
'dialog from the Taskbar, but we cant do that until
'it is fully opened. So we hook it for it's move
'event(which we call). When it moves we hide it from
'the Taskbar and then unhook - bit messy but it works.
Private Sub DialogHook()
    lpPrevWndProc = SetWindowLong(DialogWindow, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Private Sub DialogUnhook()
    SetWindowLong DialogWindow, GWL_WNDPROC, lpPrevWndProc
End Sub
Private Function WindowProc(ByVal mHwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_MOVE
            TaskbarHide
    End Select
    WindowProc = CallWindowProc(lpPrevWndProc, mHwnd, uMsg, wParam, lParam)
End Function

Public Sub SizeTV(mLeft As Long, mTop As Long, mWidth As Long, mHeight As Long)
    'Called on the resize event of the Container holding the Treeview
    Call MoveWindow(SysTreeWindow, mLeft, mTop, mWidth, mHeight, True)
End Sub

Public Sub ChangePath(mPath As String)
    'We call this sub to change the path of the Treeview
    m_CurrentDirectory = mPath 'update variable
    'Tell BrowseForFolder what to do
    Call SendMessage(DialogWindow, BFFM_SETSELECTION, 1, m_CurrentDirectory)
End Sub
