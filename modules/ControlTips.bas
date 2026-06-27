Attribute VB_Name = "ControlTips"
Option Compare Database

Public bShowing As Boolean
'Info about this at: [URL="http://support.microsoft.com/kb/152969"]http://support.microsoft.com/kb/152969[/URL]
' Access the GetCursorPos function in user32.dll
#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' GetCursorPos requires a variable declared as a custom data type
' that will hold two integers, one for x (Left)value and one for y (top)value
#End If
Type POINTAPI
    X_Pos As Long
    Y_Pos As Long
End Type

Sub GetCursorPosition(x As Long, Y As Long)
    Dim MPos As POINTAPI

    'use the API to return our public type
    GetCursorPos MPos

    'return them back outside the sub as the passed in parameters
    x = MPos.X_Pos
    Y = MPos.Y_Pos

End Sub

Public Function ShowControlTip(LeftX As Long, TopY As Long, FormNameX As String, Optional MessageX As String)
    If bShowing = False Then
        bShowing = True

        DoCmd.OpenForm FormNameX, , , , , , MessageX
    End If
    Forms(FormNameX).Move Left:=(LeftX * 10), Top:=(TopY * 10)
End Function
Public Function HideControlTip()

    DoCmd.close acForm, "ControlTipPop"
    bShowing = False
End Function
