Attribute VB_Name = "Security"
Option Compare Database
Option Explicit

Public Function Encrypt(strvalue As String) As String

Const LowerAlpha    As String = "abcdefghijklmnopqrstuvwxyz"
Const LowerSub      As String = "zebrascdfghijklmnopqtuvwxy" 'zebras
Const UpperAlpha    As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const UpperSub      As String = "ZEBRASCDFGHIJKLMNOPQTUVWXY" 'ZEBRAS

Dim lngi            As Long
Dim lngE            As Long
Dim strEncrypt      As String
Dim strLetter       As String

If strvalue & "" = "" Then Exit Function

For lngi = 1 To Len(strvalue)

    strLetter = Mid(strvalue, lngi, 1)

    Select Case Asc(strLetter)

        Case 65 To 90 'Uppercase
            'Find position in alpha string
            For lngE = 1 To Len(UpperAlpha)
                If Mid(UpperAlpha, lngE, 1) = strLetter Then GoTo USub
            Next
USub:
            strEncrypt = strEncrypt & Mid(UpperSub, lngE, 1)

        Case 97 To 122 'Lowercase
            'Find position in alpha string
            For lngE = 1 To Len(LowerAlpha)
                If Mid(LowerAlpha, lngE, 1) = strLetter Then GoTo LSub
            Next
LSub:
            strEncrypt = strEncrypt & Mid(LowerSub, lngE, 1)

        Case Else 'Do not substitute

            strEncrypt = strEncrypt & strLetter

    End Select

Next

'Now pass this string through ROT13 for another tier of security

For lngi = 1 To Len(strEncrypt)
    Encrypt = Encrypt & Chr(Asc(Mid(strEncrypt, lngi, 1)) + 13)
Next

End Function
Public Function Decrypt(strvalue As String) As String

Const LowerAlpha    As String = "abcdefghijklmnopqrstuvwxyz"
Const LowerSub      As String = "zebrascdfghijklmnopqtuvwxy" 'zebras
Const UpperAlpha    As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const UpperSub      As String = "ZEBRASCDFGHIJKLMNOPQTUVWXY" 'ZEBRAS

Dim lngi            As Long
Dim lngE            As Long
Dim strDecrypt      As String
Dim strLetter       As String

If strvalue & "" = "" Then Exit Function

'Reverse the ROT13 cipher

For lngi = 1 To Len(strvalue)
    strDecrypt = strDecrypt & Chr(Asc(Mid(strvalue, lngi, 1)) - 13)
Next

'Now reverse the encryption

For lngi = 1 To Len(strDecrypt)

    strLetter = Mid(strDecrypt, lngi, 1)

    Select Case Asc(strLetter)

        Case 65 To 90 'Uppercase
            'Find position in sub string
            For lngE = 1 To Len(UpperSub)
                If Mid(UpperSub, lngE, 1) = strLetter Then GoTo USub
            Next
USub:
            Decrypt = Decrypt & Mid(UpperAlpha, lngE, 1)

        Case 97 To 122 'Lowercase
            'Find position in sub string
            For lngE = 1 To Len(LowerSub)
                If Mid(LowerSub, lngE, 1) = strLetter Then GoTo LSub
            Next
LSub:
            Decrypt = Decrypt & Mid(LowerAlpha, lngE, 1)

        Case Else 'Do not substitute

            Decrypt = Decrypt & strLetter

    End Select

Next

End Function
