Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Utilities from http://blog.nkadesign.com/microsoft-access/
' (c) Renaud Bompuis, 2008, 2013
' Licensed under the Creative Commons Attribution License
' http://creativecommons.org/licenses/by/3.0/
' http://creativecommons.org/licenses/by/3.0/legalcode
'
' Free for re-use in any application or tutorial providing clear credit
' is made about the origin of the code and a link to the site above
' is prominently displayed where end-user can access it.
'-----------------------------------------------------------------------------
Option Compare Database
Option Explicit

'==============================================================================
' API and utilities delarations
'==============================================================================

' Various useful constants
Private Const INVALID_FILENAME_CHARS = "/\:*?<>""|"
Private Const GMEM_MOVEABLE As Long = &H2&
Private Const GMEM_ZEROINIT As Long = &H40&
Private Const CF_TEXT = 1
Private Const CF_UNICODETEXT = 13
Private Const WU_LOGPIXELSX = 88
Private Const WU_LOGPIXELSY = 90
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOW = 5
Private Const TWIPSPERINCH = 1440
Private Const INVALID_FILE_ATTRIBUTES As Long = -1

' Found from http://www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_20926813.html
Private Const RES_BT_OK = 800
Private Const RES_BT_Cancel = 801
Private Const RES_BT_Abort = 802
Private Const RES_BT_Retry = 803
Private Const RES_BT_Ignore = 804
Private Const RES_BT_Yes = 805
Private Const RES_BT_No = 806
Private Const RES_BT_Close = 807
Private Const RES_BT_Help = 808
Private Const RES_BT_TryAgain = 809
Private Const RES_BT_continue = 810

' Window Styles http://msdn.microsoft.com/en-us/library/ms632600(v=vs.85).aspx
' Returned by GetWindowLong
Private Const GWL_STYLE As Integer = -16
Private Const WS_ICONIC As Long = &H20000000
Private Const WS_MINIMIZE As Long = &H20000000
Private Const WS_MAXIMIZE As Long = &H1000000
Private Const WS_VISIBLE As Long = &H10000000

' Used to ask System for the Logical pixels/inch in X & Y axis
Private Const LOGPIXELSY = 90
Private Const LOGPIXELSX = 88

' DrawText() Format Flags
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_EDITCONTROL = &H2000&
Private Const DT_NOCLIP = &H100

' Font stuff
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_PRECIS = 4
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_OUTLINE_PRECIS = 8

Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_MASK = &HF
Private Const CLIP_LH_ANGLES = 16
Private Const CLIP_TT_ALWAYS = 32
Private Const CLIP_EMBEDDED = 128

Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255

Private Const LF_FACESIZE = 32

' Used by some API calls, like GetMonitorInfo
Private Type RECT
    Left        As Long  ' x1
    Top         As Long  ' y1
    Right       As Long  ' x2
    Bottom      As Long  ' y2
End Type

' Used by GetMonitorInfo
Private Type typMonitorInfo
    cbSize      As Long
    rcMonitor   As RECT
    rcWork      As RECT
    dwFlags     As Long
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later
  Private Declare PtrSafe Function WinAPISetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
  Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
  Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
  Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As LongPtr, ByVal lpMonitorInfo As LongPtr) As LongPtr
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
  Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
  Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long  'BOOL
  Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
  Private Declare PtrSafe Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As LongPtr, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
  Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
  Private Declare PtrSafe Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As LongPtr, lpMetrics As TEXTMETRIC) As Long
  Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
  Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
  Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As LongPtr, ByVal lpStr As LongPtr, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare PtrSafe Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
  Private Declare PtrSafe Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As LongPtr) As Long
  Private Declare PtrSafe Sub CopyString Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As LongPtr, ByVal Source As LongPtr, ByVal bytes As Long)
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else ' WIN API Declarations for Office 2007
  Private Declare Function WinAPISetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
  Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
  Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As Long, ByVal lpMonitorInfo As Long) As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
  Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
  Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
  Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function CloseClipboard Lib "user32" () As Long
  Private Declare Function EmptyClipboard Lib "user32" () As Long
  Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLib As Long) As Long  'BOOL
  Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal strFilePath As String) As Long
  Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hLib As Long, ByVal ResourceID As Long, ByVal lpBuffer As String, ByVal nBufferSize As Long) As Long
  Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
  Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
  Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
  Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
  Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
  Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
  Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
  Private Declare Sub CopyString Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As Long, ByVal Source As Long, ByVal bytes As Long)
  Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'==============================================================================
' DialogBox Constants and internal fields
'==============================================================================
Private m_Buttons As VbMsgBoxStyle
Private m_Title As String
Private m_Prompt As String
Private m_ButtonDelay As Long
Private m_IsRichText As Boolean
Private m_Result As VbMsgBoxResult
Private m_PlaintextPrompt  As String
Private m_SavedTextFileFolder As String
Private m_AllowBeep As Boolean
Private m_AllowCopyToClipboard As Boolean
Private m_AllowSaveToFile As Boolean
Private m_LabelButton1 As Variant
Private m_LabelButton2 As Variant
Private m_LabelButton3  As Variant
Private m_DismissID As Variant
#If VBA7 Then
    Private m_ParenthWnd As LongPtr
#Else
    Private m_ParenthWnd As Long
#End If


'-----------------------------------------------------------------------------
' Get the dialog Result
'-----------------------------------------------------------------------------
Public Property Get Result() As VbMsgBoxResult
    Result = m_Result
End Property

'-----------------------------------------------------------------------------
' Get/Set the dialog Buttons and Style
'-----------------------------------------------------------------------------
Public Property Get Buttons() As VbMsgBoxStyle
    Buttons = m_Buttons
End Property

Public Property Let Buttons(btn As VbMsgBoxStyle)
    m_Buttons = btn
End Property

'-----------------------------------------------------------------------------
' Get/Set the dialog Prompt message
'-----------------------------------------------------------------------------
Public Property Get Prompt() As String
    Prompt = m_Prompt
End Property

Public Property Let Prompt(Message As String)
    m_Prompt = Message
End Property

'-----------------------------------------------------------------------------
' Get/Set where the text file should be saved when the user clicks the save message buttons
'-----------------------------------------------------------------------------
Public Property Get SavedTextFileFolder() As String
    SavedTextFileFolder = m_SavedTextFileFolder
End Property

Public Property Let SavedTextFileFolder(folder As String)
    m_SavedTextFileFolder = folder
End Property

'-----------------------------------------------------------------------------
' Get/Set the dialog Prompt message
'-----------------------------------------------------------------------------
Public Property Get ButtonDelay() As Long
    ButtonDelay = m_ButtonDelay
End Property

Public Property Let ButtonDelay(delay As Long)
    If delay < 0 Then delay = 0
    m_ButtonDelay = delay
End Property

'-----------------------------------------------------------------------------
' Get/Set whether Prompt is normal text or RichText (HTML)
'-----------------------------------------------------------------------------
Public Property Get IsRichText() As Boolean
    IsRichText = m_IsRichText
End Property

Public Property Let IsRichText(isrich As Boolean)
    m_IsRichText = isrich
End Property

'-----------------------------------------------------------------------------
' Get/Set whether we should play beeps or not
'-----------------------------------------------------------------------------
Public Property Get AllowBeep() As Boolean
    AllowBeep = m_AllowBeep
End Property

Public Property Let AllowBeep(allow As Boolean)
    m_AllowBeep = allow
End Property

'-----------------------------------------------------------------------------
' Get/Set whether we should allow copy to clipboard button or not
'-----------------------------------------------------------------------------
Public Property Get AllowCopyToClipboard() As Boolean
    AllowCopyToClipboard = m_AllowCopyToClipboard
End Property

Public Property Let AllowCopyToClipboard(allow As Boolean)
    m_AllowCopyToClipboard = allow
End Property

'-----------------------------------------------------------------------------
' Get/Set whether we should allow the Save to file button or not
'-----------------------------------------------------------------------------
Public Property Get AllowSaveToFile() As Boolean
    AllowSaveToFile = m_AllowSaveToFile
End Property

Public Property Let AllowSaveToFile(allow As Boolean)
    m_AllowSaveToFile = allow
End Property

'-----------------------------------------------------------------------------
' Get/Set the dialog title
'-----------------------------------------------------------------------------
Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(caption As String)
    m_Title = caption
End Property

'-----------------------------------------------------------------------------
' Get/Set the Custom Label for button 1
'-----------------------------------------------------------------------------
Public Property Get LabelButton1() As Variant
    LabelButton1 = m_LabelButton1
End Property

Public Property Let LabelButton1(caption As Variant)
    m_LabelButton1 = caption
End Property

'-----------------------------------------------------------------------------
' Get/Set the Custom Label for button 2
'-----------------------------------------------------------------------------
Public Property Get LabelButton2() As Variant
    LabelButton2 = m_LabelButton2
End Property

Public Property Let LabelButton2(caption As Variant)
    m_LabelButton2 = caption
End Property

'-----------------------------------------------------------------------------
' Get/Set the Custom Label for button 1
'-----------------------------------------------------------------------------
Public Property Get LabelButton3() As Variant
    LabelButton3 = m_LabelButton3
End Property

Public Property Let LabelButton3(caption As Variant)
    m_LabelButton3 = caption
End Property

'-----------------------------------------------------------------------------
' Unique ID that identifies the dialog to allow the user to never see it again
'-----------------------------------------------------------------------------
Public Property Get DismissID() As Variant
    DismissID = m_DismissID
End Property

Public Property Let DismissID(ID As Variant)
    m_DismissID = ID
End Property

'-----------------------------------------------------------------------------
' Get/Set the Handle of the parent window
' Returns the MS Access main window handle if this property wans't set, so
' the dialog box will center on the monitor used by the main access window
'-----------------------------------------------------------------------------
#If VBA7 Then
    Public Property Get ParenthWnd() As LongPtr
        If m_ParenthWnd = 0 Then
            ParenthWnd = Application.hWndAccessApp
        Else
            ParenthWnd = m_ParenthWnd
        End If
    End Property

    Public Property Let ParenthWnd(handle As LongPtr)
        m_ParenthWnd = handle
    End Property
#Else
    Public Property Get ParenthWnd() As Long
        If m_ParenthWnd = 0 Then
            ParenthWnd = Application.hWndAccessApp
        Else
            ParenthWnd = m_ParenthWnd
        End If
    End Property

    Public Property Let ParenthWnd(handle As Long)
        m_ParenthWnd = handle
    End Property
#End If

'==============================================================================
' Dialog Box implementation
'==============================================================================

'-----------------------------------------------------------------------------
' Open Form Event
'-----------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    txtMessage.BorderStyle = 0 'No border by default
    txtMessage.ScrollBars = 0  'No scrollbar by default
    m_Result = 0
    m_IsRichText = False
    m_SavedTextFileFolder = CurrentProject.path & "\"
    m_Buttons = (vbOKOnly + vbCritical)
    m_ButtonDelay = 2
    m_AllowCopyToClipboard = True
    m_AllowSaveToFile = True
    btDeadFocus.SetFocus
End Sub

'-----------------------------------------------------------------------------
'  Show the form Modally
'-----------------------------------------------------------------------------
Public Function ShowModal() As VbMsgBoxResult
    caption = Title & " "

    ' To size the textbox we will need to use the raw text data
    If m_IsRichText Then
        txtMessage.TextFormat = acTextFormatHTMLRichText
        m_PlaintextPrompt = Application.PlainText(Prompt)
    Else
        txtMessage.TextFormat = acTextFormatPlain
        m_PlaintextPrompt = Prompt
    End If

    ' Initialise the buttons to show
    SetUpButtons

    ' Display/hide the clipboard and save to file buttons
    btCopyToFile.Visible = AllowSaveToFile
    btCopy.Visible = AllowCopyToClipboard

    ' Display the dismiss checkbox or not
    ckDismiss = False
    If IsBlank(DismissID) Or (Buttons And 7) <> 0 Then
        ckDismiss.Visible = False
    Else
        ckDismiss.Visible = Not GetDismissIDValue(DismissID)
    End If

    ' Calulate the proper dimensions of the textbox and the form
    Redimension

    ' Set the message
    txtMessage.Value = Prompt

    'Remove the selection on the textbox if necessary
    If txtMessage.Enabled Then txtMessage.SelLength = 0

    ' Beep - The value passed to MessageBeep corresponds to the style of the
    ' icon to be displayed and is encoded into the high nibble of the LSB
    If m_AllowBeep Then MessageBeep Buttons And &H70

    ' Make the form visible
    SetFocus

    ' Setup button delay
    TimerInterval = m_ButtonDelay * 1000
    If TimerInterval = 0 Then EnableButtons

    ' Wait until the form becomes hidden (user clicked a button)
    On Error GoTo Forced_Closed
    m_Result = -1
    Do While m_Result = -1
        DoEvents
        Sleep 50
    Loop

    ' Return the result
Forced_Closed:
    ShowModal = m_Result
End Function

'-----------------------------------------------------------------------------
'  Setup the buttons
'-----------------------------------------------------------------------------
Private Sub SetUpButtons()
    ' Detect which icon to display
    If (m_Buttons And vbExclamation) = vbExclamation Then
        imgQuestion.Visible = False
        imgInformation.Visible = False
        imgCritical.Visible = False
        imgExclamation.Visible = True
    ElseIf (m_Buttons And vbQuestion) = vbQuestion Then
        imgCritical.Visible = False
        imgExclamation.Visible = False
        imgInformation.Visible = False
        imgQuestion.Visible = True
    ElseIf (m_Buttons And vbCritical) = vbCritical Then
        imgExclamation.Visible = False
        imgQuestion.Visible = False
        imgInformation.Visible = False
        imgCritical.Visible = True
    Else
        imgQuestion.Visible = False
        imgCritical.Visible = False
        imgExclamation.Visible = False
        imgInformation.Visible = True
    End If

    ' Detect which buttons to display
    If (m_Buttons And vbCustom) = vbCustom Then
        bt3.Tag = vbBt3
        bt3.caption = Nz(m_LabelButton3)
        bt3.Visible = Not IsBlank(m_LabelButton3)
        bt2.Tag = vbBt2
        bt2.caption = Nz(m_LabelButton2)
        bt2.Visible = Not IsBlank(m_LabelButton2)
        bt1.Tag = vbBt1
        bt1.caption = Nz(m_LabelButton1)
        bt1.Visible = Not IsBlank(m_LabelButton1)
        ' In case no button has any label, fall back on vbOK only
        If Not (bt1.Visible Or bt2.Visible Or bt3.Visible) Then
            bt1.Visible = True
            bt1.Tag = vbOK
            bt1.caption = GetUser32ResourceString(RES_BT_OK)
        End If
        bt1.Cancel = True
    ElseIf (m_Buttons And vbRetryCancel) = vbRetryCancel Then
        bt3.Visible = False
        bt2.Visible = True
        bt2.Tag = vbRetry
        bt2.caption = GetUser32ResourceString(RES_BT_Retry)
        bt1.Visible = True
        bt1.Tag = vbCancel
        bt1.caption = GetUser32ResourceString(RES_BT_Cancel)
        bt1.Cancel = True
    ElseIf (m_Buttons And vbYesNo) = vbYesNo Then
        bt3.Visible = False
        bt2.Visible = True
        bt2.Tag = vbYes
        bt2.caption = GetUser32ResourceString(RES_BT_Yes)
        bt1.Visible = True
        bt1.Tag = vbNo
        bt1.caption = GetUser32ResourceString(RES_BT_No)
        bt1.Cancel = True
    ElseIf (m_Buttons And vbYesNoCancel) = vbYesNoCancel Then
        bt3.Visible = True
        bt3.Tag = vbYes
        bt3.caption = GetUser32ResourceString(RES_BT_Yes)
        bt2.Visible = True
        bt2.Tag = vbNo
        bt2.caption = GetUser32ResourceString(RES_BT_No)
        bt1.Visible = True
        bt1.Tag = vbCancel
        bt1.caption = GetUser32ResourceString(RES_BT_Cancel)
        bt1.Cancel = True
    ElseIf (m_Buttons And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
        bt3.Visible = True
        bt3.Tag = vbAbort
        bt3.caption = GetUser32ResourceString(RES_BT_Abort)
        bt2.Visible = True
        bt2.Tag = vbRetry
        bt2.caption = GetUser32ResourceString(RES_BT_Retry)
        bt1.Visible = True
        bt1.Tag = vbIgnore
        bt1.caption = GetUser32ResourceString(RES_BT_Ignore)
        bt1.Cancel = True
    ElseIf (m_Buttons And vbOKCancel) = vbOKCancel Then
        bt3.Visible = False
        bt2.Visible = True
        bt2.Tag = vbOK
        bt2.caption = GetUser32ResourceString(RES_BT_OK)
        bt1.Visible = True
        bt1.Tag = vbCancel
        bt1.caption = GetUser32ResourceString(RES_BT_Cancel)
        bt1.Cancel = True
    Else
        bt3.Visible = False
        bt2.Visible = False
        bt1.Visible = True
        bt1.Tag = vbOK
        bt1.caption = GetUser32ResourceString(RES_BT_OK)
        bt1.Cancel = True
    End If
End Sub

'-----------------------------------------------------------------------------
' Redimension the dialog to display the message in the best possible way
'-----------------------------------------------------------------------------
Private Sub Redimension()
    Dim lngWidth As Long, lngHeight As Long, lngTotalLines As Long, lngRet As Long
    Dim formWidth As Long, formHeight As Long
    Dim maxWidth As Long, maxHeight As Long
    Dim monitorWidth As Long, monitorHeight As Long
    Dim formAllMarginsHeight As Long, formAllMarginsWidth As Long
    Dim boxHeight As Long, boxWidth As Long
    Dim paddingWidth As Long, paddingHeight As Long
    Dim overflow As Boolean

    ' Get the size of the monitor where the dialog is supposed to be displayed
    #If VBA7 Then
        Dim hMonitor    As LongPtr
    #Else
        Dim hMonitor    As Long
    #End If
    Dim MonitorInfo     As typMonitorInfo

    hMonitor = MonitorFromWindow(ParenthWnd, &H0&)
    MonitorInfo.cbSize = LenB(MonitorInfo)
    GetMonitorInfo hMonitor, VarPtr(MonitorInfo)

    With MonitorInfo.rcWork
        monitorWidth = Abs(.Right - .Left)
        monitorHeight = Abs(.Bottom - .Top)
    End With

    ' Compute maximal acceptable dialog box size in twips
    monitorWidth = ConvertPixelsToTwips(monitorWidth, 0)
    monitorHeight = ConvertPixelsToTwips(monitorHeight, 0)
    maxWidth = monitorWidth * 0.6
    maxHeight = monitorHeight * 0.9

    ' Calculate the height and width of the area around the textbox
    formAllMarginsHeight = Me.WindowHeight - Section(acDetail).Height + txtMessage.Top + IIf(ckDismiss.Visible, 450, 0)
    formAllMarginsWidth = Me.Width - txtMessage.Width
    paddingWidth = txtMessage.LeftPadding + txtMessage.RightPadding
    paddingHeight = txtMessage.TopPadding + txtMessage.BottomPadding

    ' Make some adjustment to make the box a bit bigger for RichText
    ' The GetTextMetrics function uses the font information from the textbox.
    ' This gives us a better chance to have eveything fit in it.
    If m_IsRichText Then
        txtMessage.FontSize = txtMessage.FontSize + 1
    End If
    ' Compute the width of the textbox for the message we need to display
    TextWidth txtMessage, m_PlaintextPrompt, lngHeight, lngWidth
    boxWidth = lngWidth + paddingWidth
    If boxWidth > (maxWidth - formAllMarginsWidth) Then
        boxWidth = (maxWidth - formAllMarginsWidth)
        lngWidth = boxWidth - paddingWidth
    End If
    If boxWidth > txtMessage.Width Then txtMessage.Width = boxWidth

    ' Now calculate the height of the textbox for the message we need to display
    TextHeight txtMessage, m_PlaintextPrompt, lngHeight, lngWidth
    ' Adjust height for a better fit
    If m_IsRichText Then
        boxHeight = lngHeight * 1.15 + paddingHeight ' The 1.15 factor is empirical
    Else
        boxHeight = lngHeight + paddingHeight
    End If
    If boxHeight > (maxHeight - formAllMarginsHeight) Then
        overflow = True
        boxHeight = (maxHeight - formAllMarginsHeight)
    End If
    txtMessage.Height = boxHeight

    ' Make the textbox Font size normal again if we're dealing with RichText
    If m_IsRichText Then
        txtMessage.FontSize = txtMessage.FontSize - 1
    End If

    ' Assess proper width and height of the overall dialog box
    formWidth = txtMessage.Width + formAllMarginsWidth
    formHeight = txtMessage.Height + formAllMarginsHeight

    ' Adjust position of the box to the middle if there is not much text.
    If formHeight < Me.WindowHeight Then
        formHeight = Me.WindowHeight
        If txtMessage.Height < Section(acDetail).Height Then
            txtMessage.Top = (Section(acDetail).Height - txtMessage.Height) \ 2
        End If
    End If
    ' Adjust visible cues to ensure that we can still read the text when it overflows the maximum box size
    'If formWidth < Me.WindowWidth Then formWidth = Me.WindowWidth
    If overflow Then
        'reduce the height of the textbox just to make sure we see its border
        txtMessage.Height = txtMessage.Height - 5
        txtMessage.BackColor = Section("Detail").BackColor
        txtMessage.Enabled = True
        txtMessage.BorderStyle = 1 'show border
        txtMessage.ScrollBars = 2  'show scrollbar
        txtMessage.SetFocus
    End If

    ' At this point, we know how to place the dialog on the monitor
    ' However, we can't use DoCmd.MoveSize to place the window as it relies on Twips that are limited
    ' to integers, so they will overflow on a large screen. We also have the issue that DoCmd.MoveSize
    ' will place the window on the main monitor, not where the parent form or the Access app is.
    ' Here we need to use API functions, and they use pixels not twips

     Dim screenX As Long, screenY As Long
     formWidth = ConvertTwipsToPixels(formWidth, 0)
     formHeight = ConvertTwipsToPixels(formHeight, 0)
     With MonitorInfo.rcWork
        screenX = .Left + (Abs(.Right - .Left) - formWidth) \ 2
        screenY = .Top + (Abs(.Bottom - .Top) - formHeight) \ 2
     End With

     ' http://msdn.microsoft.com/en-us/library/ms633545.aspx
     SetWindowPos Me.hWnd, 0, screenX, screenY, formWidth, formHeight, &H20
End Sub


'-----------------------------------------------------------------------------
' Enable the buttons and set the focus on the default button
' This is usually called from the Form_Timer event handler after the button
' delay has elapsed.
'-----------------------------------------------------------------------------
Private Sub EnableButtons()
    bt1.Enabled = True
    bt2.Enabled = True
    bt3.Enabled = True

    Dim defaultbt As CommandButton
    Set defaultbt = bt1

    ' Detect Default button
    If (m_Buttons And vbYesNoCancel) = vbYesNoCancel Or (m_Buttons And vbAbortRetryIgnore) = vbAbortRetryIgnore Then ' 3 buttons
        If (m_Buttons And vbDefaultButton2) = vbDefaultButton2 Then
            Set defaultbt = bt2
        ElseIf (m_Buttons And vbDefaultButton3) = 0 Then
            Set defaultbt = bt3
        End If
    ElseIf (m_Buttons And vbYesNo) = vbYesNo Or (m_Buttons And vbOKCancel) = vbOKCancel Or (m_Buttons And vbRetryCancel) = vbRetryCancel Then ' 2 buttons
        If (m_Buttons And vbDefaultButton2) = 0 Then
            Set defaultbt = bt2
        End If
    End If

    ' On Error Resume Next ' Don't panic if we can't set the focus properly, the user may have badly used custom buttons
    defaultbt.Default = True
    defaultbt.SetFocus

    ' Ensures that the window is activated so that the default button receives the keyboard focus
    WinAPISetFocus Me.hWnd

    Set defaultbt = Nothing
End Sub

'-----------------------------------------------------------------------------
' Prevent the form from closing until the user has clicked one of the buttons
'-----------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    Cancel = (m_Result < 0)
End Sub

'-----------------------------------------------------------------------------
' Right button clicked
'-----------------------------------------------------------------------------
Private Sub bt1_Click()
    m_Result = bt1.Tag
    Visible = False
End Sub

'-----------------------------------------------------------------------------
' Middle button clicked
'-----------------------------------------------------------------------------
Private Sub bt2_Click()
    m_Result = bt2.Tag
    Visible = False
End Sub

'-----------------------------------------------------------------------------
' Left button clicked
'-----------------------------------------------------------------------------
Private Sub bt3_Click()
    m_Result = bt3.Tag
    Visible = False
End Sub

'-----------------------------------------------------------------------------
' Enable the buttons after the elapsed time set by ButtonDelay
'-----------------------------------------------------------------------------
Private Sub Form_Timer()
    TimerInterval = 0
    EnableButtons
End Sub

'-----------------------------------------------------------------------------
' Copy the message to the clipboard
'-----------------------------------------------------------------------------
Private Sub btCopy_Click()
    btDeadFocus.SetFocus
    ClipBoard_SetText m_PlaintextPrompt & ""
End Sub

'-----------------------------------------------------------------------------
' Copy the message to a text file.
' Here we just copy the message to the application folder
' In real life, we would do something more friendly such as create the file
' in a fodler on the user desktop or in My Documents
' Saves the file as UTF-16 with a leading BOM
'-----------------------------------------------------------------------------
Private Sub btCopyToFile_Click()
    btDeadFocus.SetFocus
    Dim intFile As Integer
    Dim fname As String
    Dim buffer() As Byte
    buffer = m_PlaintextPrompt

    fname = MakeFriendlyFileName(m_SavedTextFileFolder, "Message - " & m_Title, "txt")

    intFile = FreeFile()
    Open fname For Binary As #intFile
    Put #intFile, , &HFEFF ' Unicode BOM
    Put #intFile, , buffer
    Close #intFile

    ' Now open the saved message
    ShellExecute 0, "open", ChrW$(34) & fname & ChrW$(34), 0, 0, SW_SHOWNORMAL
End Sub

'-----------------------------------------------------------------------------
' Allow the dialog to never been displayed again
' The message is only visible if the dialog was open with a DismissID parameter.
'-----------------------------------------------------------------------------
Private Sub ckDismiss_AfterUpdate()
    ' Record the user's choice in the registry
    SetDismissIDValue DismissID, ckDismiss.Value
End Sub


'=============================================================================
' Helper functions
'=============================================================================

'-----------------------------------------------------------------------------
' Returns true if the file exists
' Use the GetFileAttributes API call instead of Dir() to avoid issue
' highlighted here:
' http://blog.nkadesign.com/2008/ms-access-enhanced-message-box-replacement/#comment-105339
'-----------------------------------------------------------------------------
'Public Function FileExists(fname As String) As Boolean
'    If IsNull(fname) Or IsEmpty(fname) Then Exit Function
'    ' Make sure that we can take care of paths longer than 260 characters
'    If Left$(fname, 2) = "\\" Then
'        FileExists = GetFileAttributes(StrPtr("\\?\UNC" & Mid$(fname, 2))) <> INVALID_FILE_ATTRIBUTES
'    Else
'        FileExists = GetFileAttributes(StrPtr("\\?\" & fname)) <> INVALID_FILE_ATTRIBUTES
'    End If
'End Function

'-----------------------------------------------------------------------------
' Returns a file name based on the given friendly name and file extension
' Usage:
'   MakeFileName("Part List", "xls")
' will return
'   [application path]\Part List 2007-09-07 15h34.txt
' where [application path] is the path to the current database.
'-----------------------------------------------------------------------------
Private Function MakeFriendlyFileName(ByVal folderPath As String, ByVal friendlyname As String, ByVal Extension As String) As String
    Dim path As String
    Dim fileName As String
    Dim AppTitle As String
    Dim fileexistcounter As Long

    ' Make sure we've got a sanitized filename and extension
    friendlyname = FilenameSanitize(Trim(friendlyname))
    If friendlyname = "" Then friendlyname = "(No filename given)"
    friendlyname = Mid$(friendlyname, 1, 32)
    Extension = Trim$(Extension)
    If Extension = "" Then Extension = "unknown"
    If Mid$(Extension, 1) <> "." Then Extension = "." & Extension

    ' Construct the path without the extension
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    path = folderPath & friendlyname & " " & Format(Now, "yyyy-MM-dd Hh" & ChrW$(34) & "h" & ChrW$(34) & "Nn")

    ' If an other file already exists, then increment the file counter until we find one that is free
    Do
        If fileexistcounter > 0 Then
            fileName = path & "(" & CStr(fileexistcounter) & ")" & Extension
        Else
            fileName = path & Extension
        End If
        fileexistcounter = fileexistcounter + 1
    Loop While FileExists(fileName)

    ' return the final filen name
    MakeFriendlyFileName = fileName

End Function

'-----------------------------------------------------------------------------
' Sanitize the given filename to ensure it is valid.
' Return a version of the input string without any invalid characters.
' Note that only a filename without path information is expected (removes the \ as well).
'-----------------------------------------------------------------------------
Private Function FilenameSanitize(fileName As String) As String
    If fileName = "" Then Exit Function

    Dim i As Integer
    Dim c As String
    Dim charlen As Long
    charlen = Len(INVALID_FILENAME_CHARS)

    For i = 1 To charlen
        c = Mid$(INVALID_FILENAME_CHARS, i, 1)
        fileName = Replace$(fileName, c, "_")
    Next i
    FilenameSanitize = fileName
End Function


'-----------------------------------------------------------------------------
' Pixel to Twips conversions
'-----------------------------------------------------------------------------
' cf http://support.microsoft.com/default.aspx?scid=kb;en-us;210590
' To call this function, pass the number of twips you want to convert,
' and another parameter indicating the horizontal or vertical measurement
' (0 for horizontal, non-zero for vertical). The following is a sample call:
'

Function ConvertTwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
   'Handle to device
   #If VBA7 Then
       Dim lngDC As LongPtr
   #Else
       Dim lngDC As Long
   #End If

   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)

   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   End If
   lngDC = ReleaseDC(0, lngDC)
   ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch
End Function

Function ConvertPixelsToTwips(lngPixels As Long, lngDirection As Long) As Long
   'Handle to device
   #If VBA7 Then
       Dim lngDC As LongPtr
   #Else
       Dim lngDC As Long
   #End If
   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)

   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   End If
   lngDC = ReleaseDC(0, lngDC)
   ConvertPixelsToTwips = (lngPixels * nTwipsPerInch) / lngPixelsPerInch
End Function


'-----------------------------------------------------------------------------
' Copy text to clipboard
' Supports strings containing Unicode characters.
'-----------------------------------------------------------------------------
Sub ClipBoard_SetText(str As String)
    If Len(str) = 0 Then Exit Sub
    #If VBA7 Then
      Dim hMem As LongPtr, pMem As LongPtr
    #Else
      Dim hMem As Long, pMem As Long
    #End If
    ' Allocate moveable global memory.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(str) + 2)

    ' Lock the block to get a far pointer
    ' to this memory.
    pMem = GlobalLock(hMem)

    ' Copy the string to this global memory
    CopyString pMem, StrPtr(str), LenB(str) + 2

    ' Unlock the memory and then copy to the clipboard
    If GlobalUnlock(hMem) = 0 Then
        If OpenClipboard(0&) <> 0 Then
            EmptyClipboard
            SetClipboardData CF_UNICODETEXT, hMem
            CloseClipboard
        End If
    End If
End Sub

'-----------------------------------------------------------------------------
' Get the given system resource string
'-----------------------------------------------------------------------------
' Inspired from http://www.thevbzone.com/l_res.htm
Private Function GetSystemResourceString(ResourcePath As String, resID As Long) As String

    #If VBA7 Then
        Dim hLibrary As LongPtr
    #Else
        Dim hLibrary As Long
    #End If
    Dim strString As String
    Dim lngStringLen As Long

    ' Load the Resource DLL
    hLibrary = LoadLibrary(ResourcePath & ChrW$(0))
    If hLibrary = 0 Then
        MsgBox "Failed to load the specified library with error code " & Err.LastDllError
        Exit Function
    End If

    ' Get a string from the Resource DLL
    strString = String$(256, 0)
    lngStringLen = LoadString(hLibrary, resID, strString, Len(strString))

    ' Close the Resource DLL
    FreeLibrary hLibrary

    GetSystemResourceString = strString
End Function

Private Function GetUser32ResourceString(resID As Long) As String
    GetUser32ResourceString = GetSystemResourceString("user32.dll", resID)
End Function


'-----------------------------------------------------------------------------
' Get the size of a string on the screen
' Note that this is a modified version of the following from Stephen Lebans' website
' http://www.lebans.com/textwidth-height.htm
' See use in the Form_FormDialog code's OnLoad event
'-----------------------------------------------------------------------------

Private Function TextHeight(ctl As Control, _
                            Optional ByVal Stext As String = "", _
                            Optional HeightTwips As Long = 0, Optional WidthTwips As Long = 0, _
                            Optional TotalLines As Long = 0, _
                            Optional TwipsPerPixel As Long = 0) As Long

    On Error Resume Next

    ' Call our function to calculate TextHeight
    ' If blWH=TRUE then we are TextHeight
    TextHeight = TextWidthOrHeight(ctl, True, Stext, HeightTwips, WidthTwips, TotalLines, TwipsPerPixel)

End Function


Private Function TextWidth(ctl As Control, _
                           Optional ByVal Stext As String = "", _
                           Optional HeightTwips As Long = 0, Optional WidthTwips As Long = 0, _
                           Optional TotalLines As Long = 0, _
                           Optional TwipsPerPixel As Long = 0) As Long

    On Error Resume Next

    ' If blWH=FALSE then we are TextWidth
    ' Call our function to calculate TextWidth
    TextWidth = TextWidthOrHeight(ctl, False, Stext, HeightTwips, WidthTwips, TotalLines, TwipsPerPixel)
End Function



Private Function TextWidthOrHeight(ctl As Control, ByVal blWH As Boolean, _
     Optional ByVal Stext As String = "", _
     Optional HeightTwips As Long = 0, Optional WidthTwips As Long = 0, _
     Optional TotalLines As Long = 0, _
     Optional TwipsPerPixel As Long = 0) As Long

     'Name                   FUNCTION() fTextWidthOrHeight
     '
     'Purpose:               Returns the Height or Width needed to
     '                       display the contents of the Control passed
     '                       to this function. This function
     '                       uses the Control's font attributes to build
     '                       a Font for the required calculations.
     '
     '                       This function replaces the Report object's TextHeight
     '                       and TextWidth methods which only work for a single line of text.
     '                       This function works with multiple lines of text and
     '                       also with both Forms and Reports.
     '
     'Version:               4.1
     '
     'Calls:                 Text API stuff. DrawText performs the actual
     '                       calculation to determine Control Height.
     '
     'Returns:               Height or width of Control in TWIPS required
     '                       to display current contents.
     '
     'Created by:            Stephen Lebans
     '
     'Credits:               If you want some...take some.
     '
     'Date:                  May 22, 2001
     '
     'Time:                  10:10:10pm
     '
     'Feedback:              Stephen@lebans.com
     '
     'My Web Page:           www.lebans.com
     '
     'Copyright:             Lebans Holdings Ltd.
     '                       Please feel free to use this code
     '                       without restriction in any application you develop.
     '                       This code may not be resold by itself or as
     '                       part of a collection.
     '
     'What's Missing:        Let me know!
     '
     '
     '
     'Bugs:
     'None at this point.
     '
     'Enjoy
     'Stephen Lebans

     '***************Code Start***************
    #If VBA7 Then
         ' Structure for DrawText calc
         Dim sRect As RECT

         ' Reports Device Context
         Dim hdc As LongPtr

         ' Holds the current screen resolution
         Dim lngDPI As Long

         Dim newfont As LongPtr
         ' Handle to our Font Object we created.
         ' We must destroy it before exiting main function

         Dim oldfont As LongPtr
         ' Device COntext's Font we must Select back into the DC
         ' before we exit this function.

         ' Logfont struct
         Dim myfont As LOGFONT

         ' TextMetric struct
         Dim tm As TEXTMETRIC

         ' LineSpacing Amount
         Dim lngLineSpacing As Long

         ' Ttemp var
         Dim numLines As Long

         ' Temp string var for current printer name
         Dim strName As String

         ' Temp vars
         Dim sngTemp1 As Single
         Dim sngTemp2 As Single
    #Else
         ' Structure for DrawText calc
         Dim sRect As RECT

         ' Reports Device Context
         Dim hdc As Long

         ' Holds the current screen resolution
         Dim lngDPI As Long

         Dim newfont As Long
         ' Handle to our Font Object we created.
         ' We must destroy it before exiting main function

         Dim oldfont As Long
         ' Device COntext's Font we must Select back into the DC
         ' before we exit this function.

         ' Logfont struct
         Dim myfont As LOGFONT

         ' TextMetric struct
         Dim tm As TEXTMETRIC

         ' LineSpacing Amount
         Dim lngLineSpacing As Long

         ' Ttemp var
         Dim numLines As Long

         ' Temp string var for current printer name
         Dim strName As String

         ' Temp vars
         Dim sngTemp1 As Single
         Dim sngTemp2 As Single
    #End If

     On Error GoTo Err_Handler

    ' Get handle to screen Device Context
    hdc = GetDC(0&)

    ' Were we passed a valid string
    If Len(Stext & vbNullString) = 0 Then
        ' Did we get a valid control passed to us?
        'select case typeof ctl is
        Select Case ctl.ControlType
            Case acTextBox
                Stext = Nz(ctl.Value, vbNullString)
            Case acLabel, acCommandButton
                Stext = Nz(ctl.caption, vbNullString)
            Case acListBox
                Stext = Nz(ctl.ItemData(0), vbNullString)
            Case Else
                ' Fail - not a control we can work with
                TextWidthOrHeight = 0
                Exit Function
        End Select
    End If


    ' Get current device resolution
    ' blWH=TRUE then we are TextHeight
    If blWH Then
        lngDPI = GetDeviceCaps(hdc, LOGPIXELSY)
    Else
        lngDPI = GetDeviceCaps(hdc, LOGPIXELSX)
    End If

    ' Calculate TwipsPerPixel
    TwipsPerPixel = TWIPSPERINCH / lngDPI

    ' We use a negative value to signify
    ' to the CreateFont function that we want a Glyph
    ' outline of this size not a bounding box.
    ' Copy font stuff from Text Control's property sheet
    With ctl
        myfont.lfClipPrecision = CLIP_LH_ANGLES
        myfont.lfOutPrecision = OUT_TT_ONLY_PRECIS
        myfont.lfEscapement = 0
        myfont.lfFaceName = .FontName & Chr$(0)
        myfont.lfWeight = .FontWeight
        myfont.lfItalic = .FontItalic
        myfont.lfUnderline = .FontUnderline
        'Must be a negative figure for height or system will return
        'closest match on character cell not glyph
        myfont.lfHeight = (.FontSize / 72) * -lngDPI
        ' Create our temp font
        newfont = CreateFontIndirect(myfont)
    End With

    If newfont = 0 Then
        Err.Raise vbObjectError + 256, "TextWidthOrHeight", "Cannot Create Font"
    End If

    ' Select the new font into our DC.
    oldfont = SelectObject(hdc, newfont)
    ' Use DrawText to Calculate height of Rectangle required to hold
    ' the current contents of the Control passed to this function.
    With sRect
        .Left = 0
        .Top = 0
        .Bottom = 0
        ' blWH=TRUE then we are TextHeight
        If blWH Then
            If WidthTwips > 0 Then
                .Right = WidthTwips / (TWIPSPERINCH / lngDPI)
            Else
                .Right = ((ctl.Width - ctl.LeftPadding - ctl.RightPadding) / (TWIPSPERINCH / lngDPI)) - 10
            End If
        Else
            If WidthTwips > 0 Then
                .Right = WidthTwips / (TWIPSPERINCH / lngDPI)
            Else
                ' Single line TextWidth
                .Right = 32000
            End If
        End If

        ' Calculate our bounding box based on the controls current width
        DrawText hdc, StrPtr(Stext), -1, sRect, DT_CALCRECT Or DT_TOP Or DT_LEFT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_EDITCONTROL Or DT_NOCLIP

        ' Get TextMetrics. This is required to determine
        ' Text height and the amount of extra spacing between lines.
        GetTextMetrics hdc, tm

        ' Cleanup
        SelectObject hdc, oldfont
        ' Delete the Font we created
        DeleteObject newfont

        ' Release the handle to the Screen's DC
        ReleaseDC 0&, hdc

        ' Calculate how many lines we are displaying
        ' return to calling function. The GDI incorrectly
        ' calculates the bounding rectangle because
        ' of rounding errors converting to Integers.
        TotalLines = .Bottom / (tm.tmHeight + tm.tmExternalLeading)
        numLines = TotalLines

        ' Convert RECT values to TWIPS
        .Bottom = (.Bottom) * (TWIPSPERINCH / lngDPI) 'sngTemp2 ' + 20

        ' Return values in optional vars
        ' Convert RECT Pixel values to TWIPS
        HeightTwips = .Bottom '* (TWIPSPERINCH / lngDPI)
        WidthTwips = .Right * (TWIPSPERINCH / lngDPI) '(GetDeviceCaps(hDC, LOGPIXELSX)))

        ' blWH=TRUE then we are TextHeight
        If blWH Then
            TextWidthOrHeight = HeightTwips
        Else
            TextWidthOrHeight = WidthTwips
        End If
    End With

    ' Exit normally
Exit_OK:
    Exit Function

Err_Handler:
    Err.Raise Err.Source, Err.Number, Err.Description
    Resume Exit_OK
End Function
