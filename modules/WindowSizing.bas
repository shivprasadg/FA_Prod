Attribute VB_Name = "WindowSizing"
Option Compare Database   'Use database order for string comparisons
Option Explicit
'* Window Manipulation Examples version 3.1
'
'    Provides examples of how to manipulate windows from within MS Access.
'
'  This module was created by:
'
'    Peter's Software
'    10540 National Blvd. #21
'    Los Angeles, CA 90034
'    USA
'    info@peterssoftware.com
'
'  Released as freeware. Comments and suggestions appreciated.
'
'  Please include this header if you pass this module along to someone else.
'  I do not require the header included in any apps you distribute, but would
'  appreciate credit being given in your documentation.
'
Global HideShow As String
Global Const WM_SM_CXVSCROLL = 2
Global Const WM_SM_CYCAPTION = 4
Global Const WM_SM_CXBORDER = 5
Global Const WM_SM_CYBORDER = 6
Global Const WM_SM_CXDLGFRAME = 7
Global Const WM_SM_CYDLGFRAME = 8
Global Const WM_SM_CYMENU = 15
Global Const WM_SM_CXFRAME = 32
Global Const WM_SM_CYFRAME = 33

Global Const WM_LOGPIXELSX = 88
Global Const WM_LOGPIXELSY = 90

Global Const WM_SWP_DRAWFRAME = &H20
Global Const WM_SWP_HIDEWINDOW = &H80
Global Const WM_SWP_NOACTIVATE = &H10
Global Const WM_SWP_NOCOPYBITS = &H100
Global Const WM_SWP_NOMOVE = &H2
Global Const WM_SWP_NOREDRAW = &H8
Global Const WM_SWP_NOREPOSITION = &H200
Global Const WM_SWP_NOSIZE = &H1
Global Const WM_SWP_NOZORDER = &H4
Global Const WM_SWP_SHOWWINDOW = &H40

Global Const WM_SW_MAXIMIZE = 3
Global Const WM_SW_MINIMIZE = 6
Global Const WM_SW_SHOWMINIMIZED = 2
Global Const WM_SW_NORMAL = 1
Global Const WM_SW_SHOWNORMAL = 1
Global Const WM_SW_RESTORE = 9
Global Const WM_SW_HIDE = 0
Global Const WM_SW_SHOW = 5

Global Const WM_HORZRES = 8
Global Const WM_VERTRES = 10

Dim i As Integer
Global iJustOpened As Integer
Dim iFormPosition As Integer                'last form was opened at this position
Global iFormPositioningMode As Integer      'this can be 1 or 2.
Global iFormTop As Integer
Global iFormLeft As Integer


'Uncomment these lines for Access version 2.0:

'Type WM_POINTAPI
'    x As Integer
'    y As Integer
'End Type

'Type WM_RECT
'    left As Integer
'    top As Integer
'    Right As Integer
'    Bottom As Integer
'End Type

'Type WM_WINDOWPLACEMENT
'    Length As Integer
'    flags As Integer
'    showCmd As Integer
'    ptMinPosition As WM_POINTAPI
'    ptMaxPosition As WM_POINTAPI
'    rcNormalPosition As WM_RECT
'End Type

'Declare Sub WM_apiSetWindowPos Lib "User" Alias "SetWindowPos" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
'Declare Function WM_apiGetParent Lib "User" Alias "GetParent" (ByVal hwnd As Integer) As Integer
'Declare Function WM_apiShowWindow Lib "User" Alias "ShowWindow" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
'Declare Function WM_apiGetDesktopWindow Lib "User" Alias "GetDesktopWindow" () As Integer
'Declare Function WM_apiGetDC Lib "User" Alias "GetDC" (ByVal hwnd As Integer) As Integer
'Declare Function WM_apiReleaseDC Lib "User" Alias "ReleaseDC" (ByVal hwnd As Integer, ByVal hdc As Integer) As Integer
'Declare Function WM_apiGetDeviceCaps Lib "GDI" Alias "GetDeviceCaps" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
'Declare Function WM_apiGetSystemMetrics Lib "User" Alias "GetSystemMetrics" (ByVal nIndex As Integer) As Integer
'Declare Function WM_apiGetWindowPlacement Lib "User" Alias "GetWindowPlacement" (ByVal hwnd As Integer, lpwndpl As WM_WINDOWPLACEMENT) As Integer
'Declare Function WM_apiiszoomed Lib "User" Alias "IsZoomed" (ByVal hwnd As Integer) As Integer
'Declare Sub WM_apiGetClientRect Lib "User" Alias "GetClientRect" (ByVal hwnd As Integer, lpRect As WM_RECT)


'Uncomment these lines for Access versions 95 and above:

Type WM_POINTAPI
    x As Long
    Y As Long
End Type

Type WM_Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type WM_WINDOWPLACEMENT
    length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As WM_POINTAPI
    ptMaxPosition As WM_POINTAPI
    rcNormalPosition As WM_Rect
End Type
#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later
Declare PtrSafe Function WM_apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare PtrSafe Function WM_apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Declare PtrSafe Function WM_apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare PtrSafe Function WM_apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Declare PtrSafe Function WM_apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Declare PtrSafe Function WM_apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare PtrSafe Function WM_apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare PtrSafe Function WM_apiGetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Declare PtrSafe Function WM_apiGetWindowPlacement Lib "user32" Alias "GetWindowPlacement" (ByVal hWnd As Long, lpwndpl As WM_WINDOWPLACEMENT) As Long
Declare PtrSafe Function WM_apiIsZoomed Lib "user32" Alias "IsZoomed" (ByVal hWnd As Long) As Long
Declare PtrSafe Function WM_apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As Long, lpRect As WM_Rect) As Long
Declare PtrSafe Function WM_apiIsIconic Lib "user32" Alias "IsIconic" (ByVal hWnd As Long) As Long
Declare PtrSafe Function WM_apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As WM_Rect) As Long
#Else
Declare Function WM_apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function WM_apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Declare Function WM_apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function WM_apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Declare Function WM_apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Declare Function WM_apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function WM_apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function WM_apiGetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Declare Function WM_apiGetWindowPlacement Lib "user32" Alias "GetWindowPlacement" (ByVal hWnd As Long, lpwndpl As WM_WINDOWPLACEMENT) As Long
Declare Function WM_apiIsZoomed Lib "user32" Alias "IsZoomed" (ByVal hWnd As Long) As Long
Declare Function WM_apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As Long, lpRect As WM_Rect) As Long
Declare Function WM_apiIsIconic Lib "user32" Alias "IsIconic" (ByVal hWnd As Long) As Long
Declare Function WM_apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As WM_Rect) As Long
#End If

Public Function MinMax()
    If HideShow = "Hide" Then
        DoCmd.RunCommand acCmdAppMinimize
        HideShow = "Show"
        Forms!MainMenu.Form.Min_Max.caption = "Show DB"
        Dim FrmNameX(25) As String
        Dim FormX As Form
        Dim OpenForms As Object
        Dim i As Integer
        i = 1
        Set OpenForms = Application.Forms
        For Each FormX In OpenForms
            FrmNameX(i) = FormX.Name
            i = i + 1
        Next FormX
        For i = 1 To Forms.Count
            DoCmd.OpenForm FrmNameX(i)
        Next i
    Else
        DoCmd.RunCommand acCmdAppRestore
        HideShow = "Hide"
        Forms!MainMenu.Form.Min_Max.caption = "Hide DB"
    End If
End Function


Function xg_GetAccesshWnd() As Long
    'Returns a handle to the main Access window
    xg_GetAccesshWnd = Application.hWndAccessApp
End Function

Function xg_GetDisplayHeight() As Integer
    '* Returns the height of the display/desktop
    Dim DisplayHeight As Integer
    Dim DisplayWidth As Integer
    Dim hDesktopWnd As Long
    Dim hDCcaps As Long
    Dim iRtn As Integer

    '* make API calls to get desktop settings
    hDesktopWnd = WM_apiGetDesktopWindow() 'get handle to desktop
    hDCcaps = WM_apiGetDC(hDesktopWnd) 'get display context for desktop
    DisplayHeight = WM_apiGetDeviceCaps(hDCcaps, WM_VERTRES)
    DisplayWidth = WM_apiGetDeviceCaps(hDCcaps, WM_HORZRES)
    iRtn = WM_apiReleaseDC(hDesktopWnd, hDCcaps) 'release display contex

    xg_GetDisplayHeight = DisplayHeight
End Function

Function xg_GetDisplayWidth() As Integer
    '* Returns the width of the display/desktop
    Dim DisplayHeight As Integer
    Dim DisplayWidth As Integer
    Dim hDesktopWnd As Long
    Dim hDCcaps As Long
    Dim iRtn As Integer

    '* make API calls to get desktop settings
    hDesktopWnd = WM_apiGetDesktopWindow() 'get handle to desktop
    hDCcaps = WM_apiGetDC(hDesktopWnd) 'get display context for desktop
    DisplayHeight = WM_apiGetDeviceCaps(hDCcaps, WM_VERTRES)
    DisplayWidth = WM_apiGetDeviceCaps(hDCcaps, WM_HORZRES)
    iRtn = WM_apiReleaseDC(hDesktopWnd, hDCcaps) 'release display contex

    xg_GetDisplayWidth = DisplayWidth
End Function
Sub xg_MaximizeWindow(sWindow As String)
    ' Maximize the active form window, or the main Access window.
    ' Parameters: "Active" - Maximize the active form window.
    '             "Access" - Maximize the main Access window
    Dim iRtn As Integer
    Dim hWndMax As Long

    If sWindow = "Active" Then
        hWndMax = Screen.ActiveForm.hWnd
    ElseIf sWindow = "Access" Then
        hWndMax = xg_GetAccesshWnd()
    Else
        MsgBox "Invalid parameter passed to xg_MaximizeWindow = " & sWindow
        Exit Sub
    End If

    iRtn = WM_apiShowWindow(hWndMax, WM_SW_MAXIMIZE)

End Sub

Sub xg_MinimizeWindow(sWindow As String)
    ' Minimize the active form window, or the main Access window.
    ' Parameters: "Active" - Minimize the active form window.
    '             "Access" - Minimize the main Access window
    Dim iRtn As Integer
    Dim hWnd As Long

    If sWindow = "Active" Then
        hWnd = Screen.ActiveForm.hWnd
    ElseIf sWindow = "Access" Then
        hWnd = xg_GetAccesshWnd()
    Else
        MsgBox "Invalid parameter passed to xg_MinimizeWindow = " & sWindow
        Exit Sub
    End If

    iRtn = WM_apiShowWindow(hWnd, WM_SW_MINIMIZE)
End Sub

Sub xg_OpenNextForm(frmname As String)
    'Save current form position, then open frmname
    Dim wp As WM_WINDOWPLACEMENT

    wp.length = Len(wp)
    If WM_apiGetWindowPlacement(Screen.ActiveForm.hWnd, wp) Then
        iFormTop = wp.rcNormalPosition.Top
        iFormLeft = wp.rcNormalPosition.Left
    End If

    DoCmd.OpenForm frmname

Exit_xg_OpenNextForm:
    Exit Sub

Err_xg_OpenNextForm:
    MsgBox Error$
    Resume Exit_xg_OpenNextForm

End Sub

Sub xg_PositionForm(frm As Form)
    'Position active form just down and to the right of the last form opened.
    'This will provide a cascade effect.

    Dim iFormsOpen As Integer
    Dim iHtOffset As Integer
    Dim iWdOffset As Integer
    Dim wp As WM_WINDOWPLACEMENT
    Dim lWindowWd As Long
    Dim lWindowHt As Long
    Dim rectClientArea As WM_Rect
    Dim lClientRight As Long
    Dim lClientBottom As Long
    Dim iYFrame As Integer
    Dim iYCaption As Integer
    Dim hDesktopWnd As Long
    Dim hDCcaps As Long
    Dim iRtn As Integer
    Dim DisplayLogPixelsY As Integer
    Dim DisplayLogPixelsX As Integer
    Dim itwipsperpixelY As Integer
    Dim itwipsperpixelX As Integer

    On Error GoTo Err_xg_PositionForm

    '* make API calls to get desktop settings and find twips per pixel
    hDesktopWnd = WM_apiGetDesktopWindow() 'get handle to desktop
    hDCcaps = WM_apiGetDC(hDesktopWnd) 'get display context for desktop
    DisplayLogPixelsY = WM_apiGetDeviceCaps(hDCcaps, WM_LOGPIXELSY)
    DisplayLogPixelsX = WM_apiGetDeviceCaps(hDCcaps, WM_LOGPIXELSX)
    iRtn = WM_apiReleaseDC(hDesktopWnd, hDCcaps) 'release display contex
    itwipsperpixelY = 1440 / DisplayLogPixelsY
    itwipsperpixelX = 1440 / DisplayLogPixelsX
    'MsgBox "itwipsperpixelY=" & itwipsperpixelY & " itwipsperpixelX=" & itwipsperpixelX

    'get height of window frame + caption to find next window position
    iYFrame = WM_apiGetSystemMetrics(WM_SM_CYFRAME) - 3
    iYCaption = WM_apiGetSystemMetrics(WM_SM_CYCAPTION)

    'Next window is placed this many twips lower and to the right of last one
    iHtOffset = (iYFrame + iYCaption + 2) * itwipsperpixelY
    iWdOffset = (iYFrame + iYCaption + 2) * itwipsperpixelX

    If WM_apiIsZoomed(frm.hWnd) Then
        'The form window is maximized - no cascading will occur.
    Else
        'Get Height and Width of window that was just opened
        wp.length = Len(wp)
        If WM_apiGetWindowPlacement(frm.hWnd, wp) Then
            'MsgBox "In relation to the Access window client area, the form window is at:  left = " & wp.rcNormalPosition.left & " top = " & wp.rcNormalPosition.top
            lWindowWd = (wp.rcNormalPosition.Right - wp.rcNormalPosition.Left) * itwipsperpixelX
            lWindowHt = (wp.rcNormalPosition.Bottom - wp.rcNormalPosition.Top) * itwipsperpixelY
        End If

        'Get Right and Bottom of Access window client area
        Call WM_apiGetClientRect(xg_GetAccesshWnd(), rectClientArea)
        lClientRight = rectClientArea.Right * itwipsperpixelX
        lClientBottom = rectClientArea.Bottom * itwipsperpixelY

        If iFormPositioningMode = 1 Then
            iFormsOpen = Forms.Count
            iFormPosition = iFormPosition + 1

            'If window will exceed width of client area, reset iFormPosition
            If (((iFormPosition - 1) * iWdOffset) + lWindowWd + (20 * itwipsperpixelX)) > lClientRight Then
                iFormPosition = 1
            End If
            'If window will exceed height of client area, reset iFormPosition
            If (((iFormPosition - 1) * iHtOffset) + lWindowHt + (20 * itwipsperpixelY)) > lClientBottom Then
                iFormPosition = 1
            End If

            'Move form to new position
            DoCmd.MoveSize (iFormPosition - 1) * iWdOffset, (iFormPosition - 1) * iHtOffset
        Else
            'If window will exceed width of client area, reset iFormPosition
            If (((iFormLeft * itwipsperpixelX) + iWdOffset) + lWindowWd + (20 * itwipsperpixelX)) > lClientRight Then
                DoCmd.MoveSize 0, 0
            ElseIf (((iFormTop * itwipsperpixelY) + iHtOffset) + lWindowHt + (20 * itwipsperpixelY)) > lClientBottom Then
                DoCmd.MoveSize 0, 0
            Else
                DoCmd.MoveSize ((iFormLeft * itwipsperpixelX) + iWdOffset), ((iFormTop * itwipsperpixelY) + iHtOffset)
            End If
        End If
    End If

Exit_xg_PositionForm:
    Exit Sub

Err_xg_PositionForm:
    If Err = 2505 Then
        DoCmd.MoveSize 0, 0
        Resume Exit_xg_PositionForm
    Else
        MsgBox "Error " & Err & " in function xg_PositionForm - " & Error$
        Resume Exit_xg_PositionForm
    End If

End Sub

Sub xg_RestoreWindow(sWindow As String)
    ' Restore the active form window, or the main Access window.
    ' Parameters: "Active" - Restore the active form window.
    '             "Access" - Restore the main Access window
    Dim iRtn As Integer
    Dim hWndRest As Long

    If sWindow = "Active" Then
        hWndRest = Screen.ActiveForm.hWnd
    ElseIf sWindow = "Access" Then
        hWndRest = xg_GetAccesshWnd()
    Else
        MsgBox "Invalid parameter passed to xg_RestoreWindow = " & sWindow
        Exit Sub
    End If

    iRtn = WM_apiShowWindow(hWndRest, WM_SW_RESTORE)

End Sub

Sub xg_SizePopUpForm(frm As Form, sngScreenPct As Single)
    '* Size the passed pop-up form to a percentage of the whole desktop area.
    '* Put this routine in your pop-up form's OnOpen event procedure.
    '* Ex.:  xg_SizePopUpForm Me, 1    '* 1 = fill 100% of desktop, 0.9 = fill 90% of desktop
    Dim DisplayHeight As Integer
    Dim DisplayWidth As Integer
    Dim DisplayLogPixelsY As Integer
    Dim WMTwipsPerPixel As Integer
    Dim hDesktopWnd As Long
    Dim hDCcaps As Long
    Dim iRtn As Integer

    If frm.PopUp Then
    Else
        MsgBox "Subroutine xg_SizePopUpForm: The form " & frm.Name & " is not a pop-up form."
        GoTo Exit_Section
    End If

    '* make API calls to get desktop settings
    hDesktopWnd = WM_apiGetDesktopWindow() 'get handle to desktop
    hDCcaps = WM_apiGetDC(hDesktopWnd) 'get display context for desktop
    DisplayHeight = WM_apiGetDeviceCaps(hDCcaps, WM_VERTRES)
    DisplayWidth = WM_apiGetDeviceCaps(hDCcaps, WM_HORZRES)
    DisplayLogPixelsY = WM_apiGetDeviceCaps(hDCcaps, WM_LOGPIXELSY)
    iRtn = WM_apiReleaseDC(hDesktopWnd, hDCcaps) 'release display contex
    WMTwipsPerPixel = 1440 / DisplayLogPixelsY

    '* Size form to a specified percentage of screen
    DoCmd.MoveSize _
        DisplayWidth * WMTwipsPerPixel * ((1 - sngScreenPct) / 2), _
        DisplayHeight * WMTwipsPerPixel * ((1 - sngScreenPct) / 3), _
        DisplayWidth * WMTwipsPerPixel * sngScreenPct, _
        DisplayHeight * WMTwipsPerPixel * sngScreenPct

Exit_Section:
End Sub
Sub xg_SizeWindow(sWindow As String, x As Integer, Y As Integer, cx As Integer, cy As Integer)
    ' Size the active form window, or the main Access window.
    ' Parameter sWindow:  "Active" - Size the active form window.
    '                     "Access" - Size the main Access window.
    '                 x:  left boundary for window (pixels)
    '                 y:  top boundary for window
    '                cx:  width of window
    '                cy:  height of window
    Dim iRtn As Integer
    Dim hWndSize As Long
    Dim strProcName As String

    On Error Resume Next
    strProcName = "xg_SizeWindow"

    If sWindow = "Active" Then
        hWndSize = Screen.ActiveForm.hWnd
        If Err <> 0 Then
            xg_ErrorMessage strProcName & " (1)"
            GoTo Exit_Section
        End If
    ElseIf sWindow = "Access" Then
        hWndSize = xg_GetAccesshWnd()
    Else
        MsgBox "Invalid parameter passed to xg_SizeWindow = " & sWindow
        Exit Sub
    End If

    iRtn = WM_apiShowWindow(hWndSize, WM_SW_RESTORE)

    Call WM_apiSetWindowPos(hWndSize, 0, x, Y, cx, cy, WM_SWP_NOZORDER Or WM_SWP_SHOWWINDOW)

    If Err <> 0 Then
        xg_ErrorMessage strProcName & " (2)"
    End If

Exit_Section:

End Sub
Sub xg_SizeFormWindow(frm As Form, x As Integer, Y As Integer, cx As Integer, cy As Integer)
    ' Size the passed form window
    ' Parameter frm:  The form whose window you want to size
    '             x:  left boundry for window (pixels)
    '             y:  top boundry for window
    '            cx:  width of window
    '            cy:  height of window
    Dim iRtn As Integer
    Dim hWndSize As Long
    Dim strProcName As String

    On Error Resume Next
    strProcName = "xg_SizeFormWindow"

    hWndSize = frm.hWnd
    If Err <> 0 Then
        xg_ErrorMessage strProcName & " (1)"
        GoTo Exit_Section
    End If

    iRtn = WM_apiShowWindow(hWndSize, WM_SW_RESTORE)

    Call WM_apiSetWindowPos(hWndSize, 0, x, Y, cx, cy, WM_SWP_NOZORDER Or WM_SWP_SHOWWINDOW)

    If Err <> 0 Then
        xg_ErrorMessage strProcName & " (2)"
    End If

Exit_Section:

End Sub
Private Sub xg_ErrorMessage(sRoutineName As String)
    MsgBox "Error in subroutine or function" & " '" & sRoutineName & "':  " & Err & " - " & Err.Description
    Err.Clear
End Sub

Sub xg_SizeWindowByPct(sWindow As String, pctcx As Integer, pctcy As Integer)
    'size the form to a percentage of the total screen size
    ' Parameter sWindow:  "Active" - Size the active form window.
    '                     "Access" - Size the main Access window.
    '             pctcx:  width of window as a percentage of desktop (i.e. 75 = 75%)
    '             pctcy:  height of window as a percentage of desktop
    Dim x As Integer, Y As Integer, cx As Integer, cy As Integer
    Dim rectClientArea As WM_Rect
    Dim AreaHeight As Integer
    Dim AreaWidth As Integer
    Dim hDesktopWnd As Long
    Dim hDCcaps As Long
    Dim iRtn As Integer
    Dim hWndSize As Long

    If pctcx >= 0 Then
        If pctcx <= 100 Then
            'pctcx is valid
        Else
            MsgBox "Invalid 'pctcx' parameter passed to 'xg_SizeWindowByPct' subroutine."
            Exit Sub
        End If
    Else
        MsgBox "Invalid 'pctcx' parameter passed to 'xg_SizeWindowByPct' subroutine."
        Exit Sub
    End If

    If pctcy >= 0 Then
        If pctcy <= 100 Then
            'pctcy is valid
        Else
            MsgBox "Invalid 'pctcy' parameter passed to 'xg_SizeWindowByPct' subroutine."
            Exit Sub
        End If
    Else
        MsgBox "Invalid 'pctcy' parameter passed to 'xg_SizeWindowByPct' subroutine."
        Exit Sub
    End If

    If sWindow = "Active" Then
        hWndSize = Screen.ActiveForm.hWnd
        Call WM_apiGetClientRect(xg_GetAccesshWnd(), rectClientArea)
        AreaHeight = rectClientArea.Bottom
        AreaWidth = rectClientArea.Right
    ElseIf sWindow = "Access" Then
        hWndSize = xg_GetAccesshWnd()
        hDesktopWnd = WM_apiGetDesktopWindow() 'get handle to desktop
        hDCcaps = WM_apiGetDC(hDesktopWnd) 'get display context for desktop
        AreaHeight = WM_apiGetDeviceCaps(hDCcaps, WM_VERTRES)
        AreaWidth = WM_apiGetDeviceCaps(hDCcaps, WM_HORZRES)
        iRtn = WM_apiReleaseDC(hDesktopWnd, hDCcaps) 'release display contex
    Else
        MsgBox "Invalid parameter passed to xg_SizeWindowByPct = " & sWindow
        Exit Sub
    End If

    iRtn = WM_apiShowWindow(hWndSize, WM_SW_RESTORE)

    cx = Int((pctcx / 100) * AreaWidth)
    cy = Int((pctcy / 100) * AreaHeight)
    x = Int((AreaWidth / 2) - (cx / 2))
    Y = Int((AreaHeight / 2) - (cy / 2))

    Call WM_apiSetWindowPos(hWndSize, 0, x, Y, cx, cy, WM_SWP_NOZORDER Or WM_SWP_SHOWWINDOW)

End Sub
