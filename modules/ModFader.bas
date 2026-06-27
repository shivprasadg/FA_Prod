Attribute VB_Name = "ModFader"
Option Compare Database
Option Explicit
#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later
Declare PtrSafe Function LoadCursorFromFileA Lib "user32" (ByVal lpFileName As String) As Long
Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#Else
Declare Function LoadCursorFromFileA Lib "user32" (ByVal lpFileName As String) As Long
Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If

Function fFadeControls(blnFadeIn As Boolean, objForm As Object, strTagSet As String)
    '====================================================================
    'COPYRIGHT (C) 2009 OPENGATE SOFTWARE, YOU ARE FREE TO USE IT IN ANY
    'OF YOUR SOFTWARE PROVIDED THE COPYRIGHT NOTICE IS NOT REMOVED OR ALTERED.
    '
    'THIS FUNCTION WILL FADE ONE OR MORE CONTROLS IN OR OUT (blnFadeIn)
    'FOR A SPECIFIED FORM (objForm) WHERE THE CONTROLS HAVE THE TAG
    'PROPERTY EQUAL TO strTagSet
    '
    'YOU CAN USE THIS ON COMBO BOXES, TEXT BOXES, AND LABELS
    '====================================================================
    Dim dblCounter  As Double
    Dim iProduct    As Integer
    Dim ctls        As Controls
    Dim ctl         As Control
    Dim dblSpeed    As Double
    Dim strParentForm As String

    On Error GoTo err_handle

    'set dblSpeed = 2 or higher if you feel
    'the fade effect is too slow
    dblSpeed = 1

    strParentForm = fGetTopForm(objForm)

    If blnFadeIn = True Then
        dblSpeed = dblSpeed * 1.5 'because there is a slight jitter to the fade in generally
        dblCounter = 255

        'fade controls from white to gray
        Do Until dblCounter < 70 'set this to 0 if you want to fade to black instead
            Set ctls = objForm.Controls
            For Each ctl In ctls
                If ctl.Tag = strTagSet Then
                    ctl.Visible = True
                    ctl.ForeColor = RGB(dblCounter, dblCounter, dblCounter)
                    ctl.BorderColor = RGB(dblCounter, dblCounter, dblCounter) 'if there is a border for a text box
                    DoCmd.RepaintObject acForm, strParentForm
                End If
            Next
            dblCounter = dblCounter - dblSpeed
        Loop


    Else

        dblCounter = 0
        'fade text from gray to white
        Do Until dblCounter > 254

            Set ctls = objForm.Controls
            For Each ctl In ctls
                If ctl.Tag = strTagSet Then
                    ctl.BackColor = RGB(dblCounter, dblCounter, dblCounter)
                    ctl.ForeColor = RGB(dblCounter, dblCounter, dblCounter) 'if there is a border for a text box
                    ctl.BorderColor = RGB(dblCounter, dblCounter, dblCounter)
                    DoCmd.RepaintObject acForm, strParentForm
                End If
            Next
            dblCounter = dblCounter + dblSpeed
        Loop

    End If



    Exit Function
err_handle:
    Select Case Err.Number
        Case 2450 'no such form exists or is loaded
            Exit Function

        Case Else
            MsgBox "Unexpected error: " & Err.Description & vbCrLf & Err.Number & _
                vbCrLf & "(line " & Erl & ") from fFadeControls"
    End Select


End Function
Function fTest()


    Forms!frmExampleUse_safe_to_delete!lblShowFader.ForeColor = RGB(255, 255, 255)

End Function
Sub sTempPause()
    Dim iCounter As Integer
    Dim iTotal   As Integer 'for compatibility with Access 2007, which is much slower on loops
    On Error Resume Next

    'If fAccess2007Plus = True Then
    iTotal = 20
    'Else
    '  iTotal = 80
    'End If

    Do Until iCounter = iTotal
        DoEvents
        iCounter = iCounter + 1
    Loop

End Sub
Function fGetTopForm(ByVal objForm As Object) As String
    '==========================================================================
    'FUNCTION RETURNS THE HIGHEST PARENT FORM NAME
    '==========================================================================
    Dim strForm As String
    Dim tmpObj  As Form

    On Error Resume Next

    strForm = objForm.Name

restart_loop:
    Set tmpObj = objForm.Parent

    Select Case Err.Number
        Case 2452 'no parent
            fGetTopForm = strForm
        Case 0
            strForm = tmpObj.Name
            Set objForm = tmpObj
            GoTo restart_loop
        Case Else
            Debug.Print "error " & Err.Number & " " & Err.Description
    End Select

End Function

Function fAccess2007Plus() As Boolean
    On Error Resume Next
    If CInt(Left(SysCmd(acSysCmdAccessVer), 2)) >= 12 Then fAccess2007Plus = True
End Function
Public Function SetMousePointer(intPointer As Integer) As Boolean
    ' Comments : sets the mouse pointer
    ' Parameters: intPointer - 0 (Default) The shape is determined by Microsoft Access
    ' 1 Normal Select (Arrow)
    ' 3 Text Select (I-Beam)
    ' 7 Vertical Resize (Size N, S)
    ' 9 Horizontal Resize (Size E, W)
    ' 11 Busy (Hourglass)
    ' Returns : True if successful, False otherwise


    On Error GoTo err_SetMousePointer

    Application.Screen.MousePointer = intPointer

    SetMousePointer = True

exit_SetMousePointer:
    Exit Function

err_SetMousePointer:
    SetMousePointer = False
    Resume exit_SetMousePointer

End Function
