Attribute VB_Name = "PixelsToTwips"
Option Compare Database
Option Explicit
#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later
Private Declare PtrSafe Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare PtrSafe Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#Else
Private Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Private Declare Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Public Const DIRECTION_VERTICAL = 1
Public Const DIRECTION_HORIZONTAL = 0

Function fTwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
'   Function to convert Twips to pixels for the current screen resolution
'   Accepts:
'       lngTwips - the number of twips to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of pixels corresponding to the given twips
    On Error GoTo E_Handle
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = apiGetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    Else
        lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    End If
    lngDeviceHandle = apiReleaseDC(0, lngDeviceHandle)
    fTwipsToPixels = lngTwips / 1440 * lngPixelsPerInch
fExit:
    On Error Resume Next
    Exit Function
E_Handle:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
    Resume fExit
End Function

Function fPixelsToTwips(lngPixels As Long, lngDirection As Long) As Long
'   Function to convert pixels to twips for the current screen resolution
'   Accepts:
'       lngPixels - the number of pixels to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:

'       the number of twips corresponding to the given pixels
    On Error GoTo E_Handle
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = apiGetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    Else
    lngPixelsPerInch = apiGetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    End If
    lngDeviceHandle = apiReleaseDC(0, lngDeviceHandle)
    fPixelsToTwips = lngPixels * 1440 / lngPixelsPerInch
fExit:
    On Error Resume Next
    Exit Function
E_Handle:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
    Resume fExit
End Function
