Attribute VB_Name = "SetWindowSizePos"
'====================================
' Global Declarations
'====================================
'  sets both X and Y to 0, sets the width to 480, and sets the height to 640.
' Use SizeAccess 0, 0, 480, 640

Option Compare Database
Option Explicit

'NOTE: The following "Declare" statement is case sensitive.
#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later
Declare PtrSafe Sub SetWindowPos Lib "user32" (ByVal hWnd&, _
                           ByVal hWndInsertAfter&, _
                           ByVal x&, ByVal Y&, ByVal cx&, _
                           ByVal cy&, ByVal wFlags&)
#Else

Declare Sub SetWindowPos Lib "user32" (ByVal hWnd&, _
                           ByVal hWndInsertAfter&, _
                           ByVal x&, ByVal y&, ByVal cx&, _
                           ByVal cy&, ByVal wFlags&)
#End If
'Moves MS Access window to top of Z-order.
Global Const HWND_TOP = 0

'Values for wFlags.
Global Const SWP_NOZORDER = &H4      'Ignores the hWndInsertAfter.

Function SizeAccess(cx As Long, cy As Long, _
  cHeight As Long, cWidth As Long)

   Dim h As Long
   'Get handle to Microsoft Access.
   h = Application.hWndAccessApp

   'Position Microsoft Access.
   SetWindowPos h, HWND_TOP, cx, cy, cWidth, _
   cHeight, SWP_NOZORDER

End Function
