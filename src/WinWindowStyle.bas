Attribute VB_Name = "WinWindowStyle"
' various Win32/Win64 APIs for seting window position, size, and state
Option Explicit

' returns True if windows was previously not hidden, False if window was hidden
' see https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
Public Declare PtrSafe Function ShowWindow Lib "user32" _
                                        (ByVal hWnd As _
                                         Long, ByVal nCmdShow As Long) As Boolean

' immediately return after requesting change
' returns True if successfully initiated window change, does not return sucess or state
Public Declare PtrSafe Function ShowWindowAsync Lib "user32" _
                                        (ByVal hWnd As Long, _
                                         ByVal nCmdShow As Long) As Boolean

' if window is iconic (minimized) then should use SW_RESTORE instead of SW_SHOWNORMAL
Public Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long

Public Declare PtrSafe Function BringWindowToTop Lib "user32" _
                                        (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Sub SetWindowPos Lib "user32" _
                                        (ByVal hWnd As Long, _
                                         ByVal hWndInsertAfter As Long, _
                                         ByVal x As Long, _
                                         ByVal y As Long, _
                                         ByVal cx As Long, _
                                         ByVal cy As Long, _
                                         ByVal wFlags As Long)

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                                        (ByVal hWndParent As Long, _
                                         ByVal hWndChildAfter As Long, _
                                         ByVal lpszWindowClass As String, _
                                         ByVal lpszWindowTitle As String) As Long

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                        (ByVal hWnd As Long, _
                                         ByVal lpString As String, _
                                         ByVal cch As Long) As Long

Enum SetWindowPosFlag
    HWND_NOTOPMOST = -2
    HWND_TOPMOST = -1
    HWND_TOP = 0
    HWND_BOTTOM = 1

    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
End Enum


Enum ShowWindowCmd
    SW_HIDE = 0            'Hides the window and activates another window.
    SW_NORMAL = 1          'Activates and displays a window. If the window is minimized or maximized, the system restores it to its original
    SW_SHOWNORMAL = 1      '  size and position. An application should specify this flag when displaying the window for the first time.
    SW_SHOWMINIMIZED = 2   'Activates the window and displays it as a minimized window.
    SW_SHOWMAXIMIZED = 3   'Activates the window and displays it as a maximized window.
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4  'Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated.
    SW_SHOW = 5            'Activates the window and displays it in its current size and position.
    SW_MINIMIZE = 6        'Minimizes the specified window and activates the next top-level window in the Z order.
    SW_SHOWMINNOACTIVE = 7 'Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.
    SW_SHOWNA = 8          'Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated.
    SW_RESTORE = 9         'Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    SW_SHOWDEFAULT = 10    'Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
    SW_FORCEMINIMIZE = 11  'Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.
End Enum


Public Enum browserType
    noBrowser = 0
    Chrome = 1
    Edge = 2
    Firefox = 4
    Chromium = browserType.Chrome Or browserType.Edge
    AnyBrowser = browserType.Chromium Or browserType.Firefox
End Enum

Private Const chromeWindowClass As String = "Chrome_WidgetWin_1" ' same window class for Edge
Private Const firefoxWindowClass As String = "MozillaWindowClass"
Private Const chromeWindowTitle As String = " - Google Chrome"
Private Const edgeWindowTitle As String = " - Microsoft? Edge" ' ? is a zero width space "&H200B" if comparing as Unicode


' returns the full title of window
Private Function getWindowTitle(ByVal hWnd As Long) As String
    Dim strBuffer As String, titleLen As Long
    ' allocate memory to store longest expected window title, note GetWindowText will truncate title returned to this length
    strBuffer = String(1024, " ")
    ' query for title using window handle
    titleLen = GetWindowText(hWnd, strBuffer, Len(strBuffer))
    ' truncate String to actual returned length so garbage characters not processed
    getWindowTitle = left(strBuffer, titleLen)
End Function

' returns hwnd of window with matching title
' to find specific browser use expected browser title
' otherwise pass "" to match first found
Public Function findBrowserHWnd(Optional ByVal title As String = vbNullString, Optional limitTo As browserType = browserType.AnyBrowser) As Long
    Dim hWnd As Long, windowTitle As String
    
    ' Edge or Chrome
    If limitTo And browserType.Chromium Then
        hWnd = FindWindowEx(0&, 0&, chromeWindowClass, vbNullString) ' just return Chrome or Edge window handles
        Do
            ' get window title so we can see does this browser handle match our expected browser? and title?
            windowTitle = getWindowTitle(hWnd)
            ' match expected browser
            If InStr(windowTitle, title) > 0 Then
                If ((limitTo And browserType.Edge) = browserType.Edge And InStr(windowTitle, edgeWindowTitle) > 0) _
                   Or ((limitTo And browserType.Chrome) = browserType.Chrome And InStr(windowTitle, chromeWindowTitle) > 0) _
                Then
                    findBrowserHWnd = hWnd
                    Exit Function
                End If
            End If
        
            ' get next possible window handle
            hWnd = FindWindowEx(0&, hWnd, chromeWindowClass, vbNullString) ' just return Chrome or Edge window handles
        Loop Until hWnd = 0 ' no more matching windows
    End If
    
    ' Firefox
    If limitTo And browserType.Firefox Then
        hWnd = FindWindowEx(0&, 0&, firefoxWindowClass, vbNullString) ' just return Firefox window handles
        Do
            ' get window title so we can see does this browser handle match our expected browser? and title?
            windowTitle = getWindowTitle(hWnd)
            ' match expected browser
            If InStr(windowTitle, title) > 0 Then
                findBrowserHWnd = hWnd
                Exit Function
            End If
        
            ' get next possible window handle
            hWnd = FindWindowEx(0&, hWnd, firefoxWindowClass, vbNullString) ' just return Chrome or Edge window handles
        Loop Until hWnd = 0 ' no more matching windows
    End If
    
    ' not found
End Function

