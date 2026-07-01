Attribute VB_Name = "Dialog"
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
Public UserInputX As String
' Buttons in the box replacement become enabled after a 2 second delay by default
Private Const DEFAULT_BUTTON_DELAY = 0

' We add a vbCustom constant to allow custom labels on buttons instead of the standard MsgBox arguments
Public Const vbCustom = 1024

' Extend the VbMsgBoxResult so we can also return additonal values if the user has selected custom buttons
Public Enum VbMsgBoxResultEx
    vbOK = VbMsgBoxResult.vbOK
    vbCancel = VbMsgBoxResult.vbCancel
    vbAbort = VbMsgBoxResult.vbAbort
    vbRetry = VbMsgBoxResult.vbRetry
    vbIgnore = VbMsgBoxResult.vbIgnore
    vbYes = VbMsgBoxResult.vbYes
    vbNo = VbMsgBoxResult.vbNo
    vbBt1 = 101
    vbBt2 = 102
    vbBt3 = 103
End Enum

'-----------------------------------------------------------------------------
' Useful Win API calls
'-----------------------------------------------------------------------------
#If VBA7 Then
  Private Declare PtrSafe Function WinAPISetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As Long
#Else
  Private Declare Function WinAPISetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
#End If

'-----------------------------------------------------------------------------
' Internal fields used to track defaul behaviour of the dialog box
'-----------------------------------------------------------------------------
Private m_DefaultTextFileFolder As Variant
Private m_DefaultButtonDelay As Variant
Private m_DefaultBeepAllowed As Variant
Private m_DefaultSaveToFileAllowed As Variant
Private m_DefaultInputRequestAllowed As Variant
Private m_DefaultCopyToClipboardAllowed As Variant
Private m_DefaultNoStrEsc As Variant
Private m_Result As VbMsgBoxResultEx

'-----------------------------------------------------------------------------
' The time in seconds to wait until the buttons becomes active on the enhanced dialog
'-----------------------------------------------------------------------------
Public Property Get DefaultButtonDelay() As Long
    If IsEmpty(m_DefaultButtonDelay) Then
        ' Initialize the default value
        m_DefaultButtonDelay = DEFAULT_BUTTON_DELAY
    End If
    DefaultButtonDelay = CLng(m_DefaultButtonDelay)
End Property

Public Property Let DefaultButtonDelay(delay As Long)
    If delay < 0 Then delay = 0
    m_DefaultButtonDelay = delay
End Property

'-----------------------------------------------------------------------------
' The default folder where the textfile will be saved
' The default value is the path to the folder where the current database is located
'-----------------------------------------------------------------------------
Public Property Get DefaultSavedTextFileFolder() As String
    If IsEmpty(m_DefaultTextFileFolder) Then
        m_DefaultTextFileFolder = CurrentProject.path & "\"
    End If
    DefaultSavedTextFileFolder = m_DefaultTextFileFolder
End Property

Public Property Let DefaultSavedTextFileFolder(folderPath As String)
    m_DefaultTextFileFolder = folderPath
End Property

'-----------------------------------------------------------------------------
' Whether the button allowing the user to save the message to file should be
' displayed or not.
' The default is yes.
'-----------------------------------------------------------------------------
Public Property Get DefaultSaveToFileAllowed() As Boolean
    If IsEmpty(m_DefaultSaveToFileAllowed) Then
        m_DefaultSaveToFileAllowed = False
    End If
    DefaultSaveToFileAllowed = m_DefaultSaveToFileAllowed
End Property

Public Property Let DefaultSaveToFileAllowed(allow As Boolean)
    m_DefaultSaveToFileAllowed = allow
End Property
' Whether the Input Box allowing the user to submit data should be
' displayed or not.
' The default is No.
'-----------------------------------------------------------------------------
Public Property Get DefaultInputRequestAllowed() As Boolean
    If IsEmpty(m_DefaultInputRequestAllowed) Then
        m_DefaultInputRequestAllowed = False
    End If
    DefaultInputRequestAllowed = m_DefaultInputRequestAllowed
End Property
Public Property Let DefaultInputRequestAllowed(allow As Boolean)
    m_DefaultInputRequestAllowed = allow
End Property
'-----------------------------------------------------------------------------
' Whether the button allowing the user to copy the message to clipboard is
' displayed or not.
' The default is yes.
'-----------------------------------------------------------------------------
Public Property Get DefaultCopyToClipboardAllowed() As Boolean
    If IsEmpty(m_DefaultCopyToClipboardAllowed) Then m_DefaultCopyToClipboardAllowed = False
    DefaultCopyToClipboardAllowed = m_DefaultCopyToClipboardAllowed
End Property

Public Property Let DefaultCopyToClipboardAllowed(allow As Boolean)
    m_DefaultCopyToClipboardAllowed = allow
End Property

'-----------------------------------------------------------------------------
' Whether beeps are allowed or not when the message box opens
' The default is yes.
'-----------------------------------------------------------------------------
Public Property Get DefaultBeepAllowed() As Boolean
    If IsEmpty(m_DefaultBeepAllowed) Then m_DefaultBeepAllowed = True
    DefaultBeepAllowed = m_DefaultBeepAllowed
End Property

Public Property Let DefaultBeepAllowed(allow As Boolean)
    m_DefaultBeepAllowed = allow
End Property

'-----------------------------------------------------------------------------
' Whether to treat the message passed to the box as raw text that should not
' be escaped (like the default MsgBox does).
' The default is no.
'-----------------------------------------------------------------------------
Public Property Get DefaultNoStrEsc() As Boolean
    If IsEmpty(m_DefaultNoStrEsc) Then m_DefaultNoStrEsc = True
    DefaultNoStrEsc = m_DefaultNoStrEsc
End Property

Public Property Let DefaultNoStrEsc(allow As Boolean)
    m_DefaultNoStrEsc = allow
End Property

'-----------------------------------------------------------------------------
' The value of the last result returned by the dialog box
'-----------------------------------------------------------------------------
Public Property Get Result() As VbMsgBoxResultEx
    Result = m_Result
End Property

' Normally., this is not needed, but there are cases where you might want to
' manually change the visible result of the DialogBox, for instance if you use
' it in another library.
Public Property Let Result(newResult As VbMsgBoxResultEx)
    m_Result = newResult
End Property

'-----------------------------------------------------------------------------
' Plain Text Replacement for the standard MsgBox
'-----------------------------------------------------------------------------
Public Function Box(ByVal Prompt As String, _
                       Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                       Optional ByVal Title As String, _
                       Optional ByVal HelpFile As String, _
                       Optional ByVal HelpContextId As Long, _
                       Optional ByVal ButtonDelay As Long = -1, _
                       Optional ByVal AllowBeep As Variant, _
                       Optional ByVal AllowCopyToClipboard As Variant, _
                       Optional ByVal AllowSaveToFile As Variant, _
                       Optional ByVal LabelButton1 As Variant = Null, _
                       Optional ByVal LabelButton2 As Variant = Null, _
                       Optional ByVal LabelButton3 As Variant = Null, _
                       Optional ByVal NoStrEsc As Variant, _
                       Optional ByVal InputRequest As Variant, _
                       Optional ByVal DismissID As Variant = Null) _
                       As VbMsgBoxResultEx

    ' If the dialog was previously dismissed, don't display anything at all
    If (Buttons And 7) = 0 And GetDismissIDValue(DismissID) Then
        Box = VbMsgBoxResultEx.vbOK
        Exit Function
    End If

    ' Create the MessageBox
    Dim f As New Form_FormDialog
    f.ParenthWnd = GetParentWindowHandle()
    f.Title = Title
    f.Buttons = Buttons
    f.HelpFile = HelpFile
    f.HelpContextId = HelpContextId
    f.ButtonDelay = IIf(ButtonDelay < 0, DefaultButtonDelay, ButtonDelay)
    f.AllowBeep = IIf(IsMissing(AllowBeep), DefaultBeepAllowed, AllowBeep)
    f.AllowCopyToClipboard = IIf(IsMissing(AllowCopyToClipboard), DefaultCopyToClipboardAllowed, AllowCopyToClipboard)
    f.AllowSaveToFile = IIf(IsMissing(AllowSaveToFile), DefaultSaveToFileAllowed, AllowSaveToFile)
    f.SavedTextFileFolder = DefaultSavedTextFileFolder
    f.IsRichText = False
    f.LabelButton1 = LabelButton1
    f.LabelButton2 = LabelButton2
    f.LabelButton3 = LabelButton3
    f.InputRequest = IIf(IsMissing(InputRequest), DefaultInputRequestAllowed, InputRequest)
    f.DismissID = DismissID
    If IIf(IsMissing(NoStrEsc), DefaultNoStrEsc, NoStrEsc) Then
        f.Prompt = Prompt
    Else
        f.Prompt = UnEscStr(Prompt, toHtml:=False)
    End If

    ' Make it visible and wait forthe user until we get the result
    Box = f.ShowModal()

    ' Keep the last result just in case the user needs it again later
    m_Result = Box

    ' Set focus back to the parent form
    On Error Resume Next
    WinAPISetFocus f.ParenthWnd
    On Error GoTo 0

    ' cleanup
    Set f = Nothing

End Function

'-----------------------------------------------------------------------------
' Rich-Text Replacement for the standard MsgBox
'-----------------------------------------------------------------------------
Public Function RichBox(ByVal Prompt As String, _
                       Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                       Optional ByVal Title As String, _
                       Optional ByVal HelpFile As String, _
                       Optional ByVal HelpContextId As Long, _
                       Optional ByVal ButtonDelay As Long = -1, _
                       Optional ByVal AllowBeep As Variant, _
                       Optional ByVal AllowCopyToClipboard As Variant, _
                       Optional ByVal AllowSaveToFile As Variant, _
                       Optional ByVal LabelButton1 As Variant = Null, _
                       Optional ByVal LabelButton2 As Variant = Null, _
                       Optional ByVal LabelButton3 As Variant = Null, _
                       Optional ByVal NoStrEsc As Variant, _
                       Optional ByVal InputRequest As Variant, _
                       Optional ByVal DismissID As Variant = Null) _
                       As VbMsgBoxResultEx

    ' If the dialog was previously dismissed, don't display anything at all
    If (Buttons And 7) = 0 And GetDismissIDValue(DismissID) Then
        RichBox = VbMsgBoxResultEx.vbOK
        Exit Function
    End If

    ' Create the MessageBox
    Dim f As New Form_FormDialog
    f.ParenthWnd = GetParentWindowHandle()
    f.Prompt = UnEscStr(Prompt, toHtml:=True)
    f.Title = Title
    f.Buttons = Buttons
    f.HelpFile = HelpFile
    f.HelpContextId = HelpContextId
    f.ButtonDelay = IIf(ButtonDelay < 0, DefaultButtonDelay, ButtonDelay)
    f.AllowBeep = IIf(IsMissing(AllowBeep), DefaultBeepAllowed, AllowBeep)
    f.AllowCopyToClipboard = IIf(IsMissing(AllowCopyToClipboard), DefaultCopyToClipboardAllowed, AllowCopyToClipboard)
    f.AllowSaveToFile = IIf(IsMissing(AllowSaveToFile), DefaultSaveToFileAllowed, AllowSaveToFile)
    f.SavedTextFileFolder = DefaultSavedTextFileFolder
    f.IsRichText = True
    f.LabelButton1 = LabelButton1
    f.LabelButton2 = LabelButton2
    f.LabelButton3 = LabelButton3
    f.InputRequest = IIf(IsMissing(InputRequest), DefaultInputRequestAllowed, InputRequest)
    f.DismissID = DismissID
    If IIf(IsMissing(NoStrEsc), DefaultNoStrEsc, NoStrEsc) Then
        f.Prompt = Prompt
    Else
        f.Prompt = UnEscStr(Prompt, toHtml:=True)
    End If

    ' Make it visible and wait forthe user until we get the result
    RichBox = f.ShowModal()

    ' Keep the last result just in case the user needs it again later
    m_Result = RichBox

    ' Set focus back to the parent form
    On Error Resume Next
    WinAPISetFocus f.ParenthWnd
    On Error GoTo 0

    ' cleanup
    Set f = Nothing

End Function


'-----------------------------------------------------------------------------
' Return the hWnd windows handle of the calling form, or default to the main Access window
' This is used so we can properly display the dialog on the monitor that contains the
' form that opened the dialog box
'-----------------------------------------------------------------------------
#If VBA7 Then
    Private Function GetParentWindowHandle() As LongPtr
        On Error Resume Next
        GetParentWindowHandle = Screen.ActiveForm.hWnd
        On Error GoTo 0
        If GetParentWindowHandle = 0 Then GetParentWindowHandle = Application.hWndAccessApp
    End Function
#Else
    Private Function GetParentWindowHandle() As Long
        On Error Resume Next
        GetParentWindowHandle = Screen.ActiveForm.hWnd
        On Error GoTo 0
        If GetParentWindowHandle = 0 Then GetParentWindowHandle = Application.hWndAccessApp
    End Function
#End If

'=============================================================================
' The following code comes from this article:
' http://blog.nkadesign.com/2013/access-using-net-strings-in-vba-for-fun-an-profit/
' The code is used to bring support for character litterals and Unicode escape sequences
' in the string to be displayed in the message box.
' Refer to the article for all details.
'=============================================================================

'-----------------------------------------------------------------------------
' Takes a string that may contain .Net escaped characters and returns the
' un-escaped version of the string.
' Note that the only difference with the spec is that we can't escape \" in VBA
' so we ignore this particular case.
' In .Net, escaped sequences in strings are converted at compile time, so we end-up
' with an issue when trying to use our StringBuilder class or sFormat when they contain
' special escape chars such as \n or litteral unicode like \u20AC since they will not be
' converted to their string representation.
' This function solves this problem.
' See:
' - http://www.codeproject.com/Articles/371232/Escaping-in-Csharp-characters-strings-string-forma
' - http://msdn.microsoft.com/en-us/library/aa691087(v=vs.71).aspx
' - http://msdn.microsoft.com/en-us/library/aa664669(v=vs.71).aspx
'-----------------------------------------------------------------------------
Private Function UnEscStr(ByRef str As Variant, toHtml As Boolean) As String
    If VarType(str) <> vbString Then Exit Function

    ' To make conversion easy and fast, we store converted character litterals
    ' in a static array so we only need to construct this once.
    Static escVals(10) As String
    Static escValsHtml(10) As String
    Static isInitialised As Boolean
    If Not isInitialised Then
        isInitialised = True
        escVals(1) = "'"
        escVals(2) = "\"
        escVals(3) = ChrW$(0)    ' Null
        escVals(4) = ChrW$(7)    ' Alert
        escVals(5) = ChrW$(8)    ' Backspace
        escVals(6) = ChrW$(&HC)  ' Form feed
        escVals(7) = vbCrLf      ' New line
        escVals(8) = vbCr        ' Carriage return
        escVals(9) = vbTab       ' Tab
        escVals(10) = ChrW$(&HB) ' Vertical tab

        escValsHtml(1) = "'"
        escValsHtml(2) = "\"
        escValsHtml(3) = vbNullString  ' Null
        escValsHtml(4) = vbNullString  ' Alert
        escValsHtml(5) = vbNullString  ' Backspace
        escValsHtml(6) = vbNullString  ' Form feed
        escValsHtml(7) = "<br/>"       ' New line
        escValsHtml(8) = vbNullString  ' Carriage return
        escValsHtml(9) = "&nbsp;&nbsp;&nbsp;&nbsp;"  ' Tab
        escValsHtml(10) = vbNullString ' Vertical tab
    End If

    Dim Start As Long, pos As Long, i As Long, u As Long
    Dim c As String, v As String
    Dim L As Integer
    Start = 1
    ' We'll search for each '\' and
    Do
        pos = InStr(Start, str, "\", vbBinaryCompare)

        ' No '\' found, we're finished with this string
        If pos = 0 Then
            If Start = 1 Then
                ' There was nothing to un-escape at all in the original
                ' string, so we just return it verbatim.
                UnEscStr = str
            Else
                ' Append the final bit of the original string to the result
                UnEscStr = UnEscStr & Mid$(str, Start)
            End If
            Exit Do
        End If

        ' Append the bit of string we found before the escaped sequence
        UnEscStr = UnEscStr & Mid$(str, Start, pos - Start)
        pos = pos + 1
        Start = pos + 1
        c = Mid$(str, pos, 1)
        ' Search the escaped character in our list of valid ones
        i = InStr(1, "'\0abfnrtv", c, vbBinaryCompare)
        If i > 0 Then
            ' We found a standard character litteral, just append
            ' its converted version to the result
            If toHtml Then
                UnEscStr = UnEscStr & escValsHtml(i)
            Else
                UnEscStr = UnEscStr & escVals(i)
            End If
        Else
            'Deal with Unicode
            Select Case AscW(c)
                Case 85  ' U: UTF-32 defined unicode point
                    ' Must have exactly 8 digits
                    ' Get the whole character code point
                    u = -1
                    On Error Resume Next
                    u = CLng("&H" & Mid$(str, Start, 8))
                    On Error GoTo 0
                    If u = -1 Then
                        ' Conversion failed, probably \Uxxx was not a unicode escape sequence but a path, like \Users
                        ' revert and continue
                        UnEscStr = UnEscStr & "\U"
                    Else
                        ' Convert unicode code point correctly
                        If toHtml Then
                            v = "&#x" & Hex$(u) & ";"
                        Else
                            If u < &H10000 Then
                                ' Standard unicode point that fits into UTF-16
                                v = ChrW$(u)
                            ElseIf u < &H10FFFF Then
                                ' We need to convert the UTF-32 into to a UTF-16 surrogate pair
                                ' See calculation http://www.unicode.org/faq/utf_bom.html#utf16-4
                                v = ChrW$(&HD7C0& + (u \ 1024))
                                UnEscStr = UnEscStr & v
                                v = ChrW$(&HDC00& + (u And &H3FF&))
                            End If
                        End If
                        UnEscStr = UnEscStr & v
                        Start = Start + 8
                    End If

                Case 117 ' u: Standard unicode point that fits into UTF-16
                    u = -1
                    On Error Resume Next
                    ' Must have exactly 4 digits
                    u = CLng("&H" & Mid$(str, Start, 4))
                    On Error GoTo 0
                    If u = -1 Then
                        ' Conversion failed, probably \Uxxx was not a unicode escape sequence but a path, like \Users
                        ' revert and continue
                        UnEscStr = UnEscStr & "\u"
                    Else
                        If toHtml Then v = "&#x" & Hex$(u) & ";" Else v = ChrW$(u)
                        UnEscStr = UnEscStr & v
                        Start = Start + 4
                    End If

                Case 120 ' x: Standard unicode point that fits into UTF-16
                    ' Variable length, 1 to 4 digits
                    pos = Start: L = 0
                    Do While InStr(1, "0123456789ABCDEFabcdef", Mid$(str, pos, 1), vbBinaryCompare) > 0 And (L < 4)
                        L = L + 1
                        pos = pos + 1
                    Loop
                    u = CLng("&H" & Mid$(str, Start, L))
                    If toHtml Then v = "&#x" & Hex$(u) & ";" Else v = ChrW$(u)
                    UnEscStr = UnEscStr & v
                    Start = Start + L

                Case Else
                    ' Unknown escape sequence, just leave it as is in the result
                    UnEscStr = UnEscStr & "\" & c
            End Select
        End If
    Loop While True
End Function

'-----------------------------------------------------------------------------
' Escape Backslashes so they can be printed correctly in the dialog box instead
' of being interpreted as escape sequences.
'-----------------------------------------------------------------------------
Public Function EscBackslash(str As String) As String
    EscBackslash = Replace$(str, "\", "\\")
End Function

'-----------------------------------------------------------------------------
' Get whether the message with the given DismissID should be shown again or not.
' Uses the standard VBA registry function, so the data is saved under:
' HKEY_CURRENT_USER\Software\VB and VBA Program Settings\DialogBox
'-----------------------------------------------------------------------------
Public Function GetDismissIDValue(DismissID As Variant) As Boolean
    If IsBlank(DismissID) Then Exit Function
    GetDismissIDValue = CBool(GetSetting(CurrentProject.Name, "DialogBox", "DismissID_" & DismissID, False))
End Function

'-----------------------------------------------------------------------------
' Save whether the message with the given DismissID should be shown again or not.
'-----------------------------------------------------------------------------
Public Sub SetDismissIDValue(DismissID As Variant, Value As Boolean)
    If IsBlank(DismissID) Then Exit Sub
    SaveSetting CurrentProject.Name, "DialogBox", "DismissID_" & DismissID, Value
End Sub

'-----------------------------------------------------------------------------
' Reset the given DismissID to show its corresponding message.
'-----------------------------------------------------------------------------
Public Sub ResetDismissID(DismissID As Variant)
    If IsBlank(DismissID) Then Exit Sub
    SaveSetting CurrentProject.Name, "DialogBox", "DismissID_" & DismissID, False
End Sub

'-----------------------------------------------------------------------------
' Reset all the current project's DismissID so all dialogs with show again.
'-----------------------------------------------------------------------------
Public Sub ResetAllDismissID()
    DeleteSetting CurrentProject.Name, "DialogBox"
End Sub

'-----------------------------------------------------------------------------
' True if the argument is Nothing, Null, Empty, Missing or an empty string .
' http://blog.nkadesign.com/2009/access-checking-blank-variables/
'-----------------------------------------------------------------------------
Public Function IsBlank(arg As Variant) As Boolean
    Select Case VarType(arg)
        Case vbEmpty
            IsBlank = True
        Case vbNull
            IsBlank = True
        Case vbString
            IsBlank = LenB(arg) = 0
        Case vbObject
            IsBlank = (arg Is Nothing)
        Case Else
            IsBlank = IsMissing(arg)
    End Select
End Function

Public Function OpenInput(ByVal MsgText As String, MessageStyle As String, MessageTitle As String, Bt1Caption As String, Optional ReasonText As String) As String
Dim CommandText As String
CommandText = MsgText & "|" & MessageStyle & "|" & MessageTitle
If Bt1Caption Like "*date*" Then CommandText = CommandText & "|Enter Date|" & Bt1Caption Else CommandText = CommandText & "|Reason|" & Bt1Caption
If ReasonText <> "" Then CommandText = CommandText & "|" & ReasonText
    DoCmd.OpenForm "PromptUserForm", , , , , acDialog, CommandText
    If IsOpen("PromptUserForm") = False Then OpenInput = UserInputX

    'Forms!PromptUserForm.Form.caption = AtlaasVersion
    'Forms!PromptUserForm.Message = "<B>" & MessageTitle & "</b><Br />" & Message
    'If IsNull(Bt1Caption) Then Bt1Caption = "OK"
    'If MessageStyle = "Information" Then Forms!PromptUserForm.imgInformation.Visible = True
    'If MessageStyle = "Warning" Then Forms!PromptUserForm.imgWarning.Visible = True
    'If MessageStyle = "Critical" Then Forms!PromptUserForm.imgCritical.Visible = True
    'If MessageStyle = "Question" Then Forms!PromptUserForm.imgQuestion.Visible = True
    'Forms!PromptUserForm.bt1.caption = Bt1Caption
    'If IsNull(Bt2Caption) Then
    'Else
    'End If
    'If IsNull(Bt3Caption) Then
    'Else
    'End If

End Function
