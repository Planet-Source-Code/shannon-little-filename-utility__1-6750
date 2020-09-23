Attribute VB_Name = "modWinProc"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
            
Private Declare Function DefWindowProc Lib "user32" Alias _
    "DefWindowProcA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
      
Declare Function CallWindowProc Lib "user32" Alias _
    "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias _
    "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Sub CopyMemoryToMinMaxInfo Lib "KERNEL32" Alias _
    "RtlMoveMemory" ( _
    hpvDest As MINMAXINFO, _
    ByVal hpvSource As Long, _
    ByVal cbCopy As Long)
        
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "KERNEL32" Alias _
    "RtlMoveMemory" ( _
    ByVal hpvDest As Long, _
    hpvSource As MINMAXINFO, _
    ByVal cbCopy As Long)

Public Const GWL_WNDPROC = -4
'Window event constants
Private Const WM_GETMINMAXINFO = &H24       'This message is sent just before a user operation that affects the window's size.
Private Const WM_ACTIVATEAPP = &H1C         'This message is sent to the active window when the application is being activated (wParam = True) or deactivated (wParam = False).
Private Const WM_COMPACTING = &H41          'This message is sent to all top-level windows when the operating system is low on memory.
Private Const WM_DISPLAYCHANGE = &H7E       'This message is sent to all top-level windows after the display resolution has changed.
Private Const WM_MOVE = &H3                 'This message is sent to a window or a control that has been moved.
Private Const WM_MOVING = &H216             'This message is sent to a window or control while it is being moved.
        'These are used by the WM_MOVING and WM_SIZING messages
        Public Const WMSZ_BOTTOM = 6
        Public Const WMSZ_BOTTOMLEFT = 7
        Public Const WMSZ_BOTTOMRIGHT = 8
        Public Const WMSZ_LEFT = 1
        Public Const WMSZ_RIGHT = 2
        Public Const WMSZ_TOP = 3
        Public Const WMSZ_TOPLEFT = 4
        Public Const WMSZ_TOPRIGHT = 5

Private Const WM_NCHITTEST = &H84           'This message is sent to a window (or control) whenever a Mouse event occurs.
        'This are the values that can be returned by a WM_NCHITTEST message
        Public Const HTBORDER = 18
        Public Const HTBOTTOM = 15
        Public Const HTBOTTOMLEFT = 16
        Public Const HTBOTTOMRIGHT = 17
        Public Const HTCAPTION = 2
        Public Const HTCLIENT = 1
        Public Const HTCLOSE = 20
        Public Const HTERROR = (-2)
        Public Const HTGROWBOX = 4
        Public Const HTHELP = 21
        Public Const HTHSCROLL = 6
        Public Const HTLEFT = 10
        Public Const HTMAXBUTTON = 9
        Public Const HTMENU = 5
        Public Const HTMINBUTTON = 8
        Public Const HTNOWHERE = 0
        Public Const HTOBJECT = 19
        Public Const HTREDUCE = HTMINBUTTON
        Public Const HTRIGHT = 11
        Public Const HTSIZE = HTGROWBOX
        Public Const HTSIZEFIRST = HTLEFT
        Public Const HTSYSMENU = 3
        Public Const HTTOP = 12
        Public Const HTTOPLEFT = 13
        Public Const HTTOPRIGHT = 14
        Public Const HTTRANSPARENT = (-1)
        Public Const HTVSCROLL = 7
        Public Const HTZOOM = HTMAXBUTTON
        
Private Const WM_PAINT = &HF                'This message is sent to a window or control to request an update to its client area (similar to the Form_Paint event).
Private Const WM_SETCURSOR = &H20           'This message is sent to a window when a mouse action is performed while the cursor is over it or one of its child windows.
Private Const WM_SIZING = &H214             'This message is sent to a window or control when it is being resized.
'WM_WINDOWPOSCHANGED


Public IsHooked As Boolean
Private myForm As Form
Global lpPrevWndProc As Long
Global gHW As Long



'Used to set min/max limits of window
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI          'Not used
    ptMaxSize As POINTAPI           'Max size when window is maximized
    ptMaxPosition As POINTAPI       'Maximized window's position
    ptMinTrackSize As POINTAPI      'Minimum tracking width/height
    ptMaxTrackSize As POINTAPI      'Maximum track width/height
End Type

Public Sub Initialize(ByVal fForm As Form)
    'myForm = fForm
End Sub

Public Sub Hook()
    If IsHooked Then
        'Already hooked
    Else
        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
        IsHooked = True
    End If
End Sub

Public Sub Unhook()
    Dim temp As Long
    'End sub-classing
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
    IsHooked = False
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
                    ByVal wParam As Long, ByVal lParam As Long) As Long
    'Main select case statement
    'Debug.Print "Message: "; hw, uMsg, wParam, lParam
    Dim MinMax As MINMAXINFO
    
    Select Case uMsg
        Case WM_MOVE:           'Message is sent when the window is moved
        
        Case WM_GETMINMAXINFO:  'Message is sent to a window when the size or position of the window is about to change
        
            'Retrieve default MinMax settings
            CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

            'Specify new minimum size for window.
            MinMax.ptMinTrackSize.x = 480
            MinMax.ptMinTrackSize.y = 250

            'Specify new maximum size for window.
            MinMax.ptMaxTrackSize.x = 480
            MinMax.ptMaxTrackSize.y = 250

            'Copy local structure back.
            CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

            'WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    End Select
    
    'Send message on through
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function


'Snaps a window to the edge of the screen if it is dragged close enough
'In the win proc, it is diabled if a control key is being pressed
Public Sub SnapToEdge(ByVal myForm As Form)
    Dim intSnapDistance As Integer
    intSnapDistance = 500
    'Snaps a form to the edge of the screen
    'Snap to left side
    If myForm.Left < intSnapDistance Then
        myForm.Left = 0
    Else    'Snap to right side
        If (myForm.Left - myForm.Width) < intSnapDistance Then
            myForm.Left = myForm.Left - myForm.Width
        End If
    End If
    'Snap to top
    If myForm.Top < intSnapDistance Then
        myForm.Top = 0
    Else    'Snap to bottom
        If (myForm.Top - myForm.Height) < intSnapDistance Then
            myForm.Top = myForm.Top - myForm.Height
        End If
    End If
End Sub

' Sets whatever window's hWnd you pass to it to either normal or always on top
Public Function SetTopWindow(hWnd As Long, blnTopOrNormal As Boolean) As Long
    Dim SWP_NOMOVE
    Dim SWP_NOSIZE
    Dim FLAGS
    Dim HWND_TOPMOST
    Dim HWND_NOTOPMOST
    
    SWP_NOMOVE = 2
    SWP_NOSIZE = 1
    FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    
    If blnTopOrNormal = True Then 'Make the window the topmost
        SetTopWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else    'Make it normal
        SetTopWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopWindow = False
    End If
End Function
