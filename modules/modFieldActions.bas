Attribute VB_Name = "modFieldActions"
Option Compare Database
Option Explicit

Public Function SetLPAFItems(NewValue As Variant, TargetField As String, fLabel As String, uLevel As String, SpecId As Long) As Integer
    Dim askMsg1 As String, askMsg2 As String, askMsg3 As String, askMsg4 As String
    Dim SQLz As String
    Dim Committ As VbMsgBoxResult
    If Nz(NewValue, "") = "" Then Exit Function
    askMsg2 = vbNewLine & vbNewLine & "Do you want all of the " & uLevel & "s in this ATF Spec to have their "
    askMsg4 = vbNewLine & vbNewLine & "Yes or No?"

    Select Case VarType(NewValue)

        Case vbInteger, vbLong, vbDecimal
            askMsg1 = "You're about to set the " & fLabel & " of this LPAF" & vbNewLine & "To: " & NewValue & "."
            askMsg3 = vbNewLine & fLabel & " Set to: " & NewValue & "."
            If uLevel = "Unit" Then SQLz = "UPDATE Units SET " & TargetField & "=" & NewValue & " WHERE " & TargetField & "<>" & NewValue & " AND ATFUnitSpecID=" & SpecId & ";"
            If uLevel = "Group" Then SQLz = "UPDATE SCHGRP SET " & TargetField & "=" & NewValue & " WHERE " & TargetField & "<>" & NewValue & " AND ATFGroupSpecID=" & SpecId & ";"

        Case vbCurrency
            askMsg1 = "You're about to set the " & fLabel & " of this LPAF" & vbNewLine & "To: " & Format(NewValue, "$#,##0.00") & "."
            askMsg3 = vbNewLine & fLabel & " Set to: " & Format(NewValue, "$#,##0.00") & "."
            If uLevel = "Unit" Then SQLz = "UPDATE Units SET " & TargetField & "=" & NewValue & " WHERE " & TargetField & "<>" & NewValue & " AND ATFUnitSpecID=" & SpecId & ";"
            If uLevel = "Group" Then SQLz = "UPDATE SCHGRP SET " & TargetField & "=" & NewValue & " WHERE " & TargetField & "<>" & NewValue & " AND ATFGroupSpecID=" & SpecId & ";"

        Case vbString
            NewValue = Replace(NewValue, "'", "''")
            askMsg1 = "You're about to set the " & fLabel & " of this LPAF" & vbNewLine & "To: " & NewValue & "."
            askMsg3 = vbNewLine & fLabel & " Set to: " & NewValue & "."
            If uLevel = "Unit" Then SQLz = "UPDATE Units SET " & TargetField & "='" & NewValue & "' WHERE " & TargetField & "<>'" & NewValue & "' AND ATFUnitSpecID=" & SpecId & ";"
            If uLevel = "Group" Then SQLz = "UPDATE SCHGRP SET " & TargetField & "='" & NewValue & "' WHERE " & TargetField & "<>'" & NewValue & "' AND ATFGroupSpecID=" & SpecId & ";"

        Case Else
            SetLPAFItems = 2
            Exit Function
    End Select

    Committ = MsgBox(askMsg1 & askMsg2 & askMsg3 & askMsg4, vbQuestion + vbYesNo + vbCritical, "Set " & TargetField & "?")

    If Committ = vbNo Then
        SetLPAFItems = 2
        Exit Function
    End If

    On Error GoTo FixMe
    DoCmd.SetWarnings True
    DoCmd.RunSQL SQLz
    DoCmd.SetWarnings True

    SetLPAFItems = 1
    Exit Function

FixMe:
    SetLPAFItems = 2
End Function
Public Function ChangeLPAFItems(NewValue As Variant, OldValue As Variant, TargetField As String, fLabel As String, uLevel As String, SpecId As Long) As Boolean
    Dim askMsg1 As String, askMsg2 As String, askMsg3 As String, askMsg4 As String
    Dim SQLz As String
    Dim Committ As VbMsgBoxResult

    askMsg2 = vbNewLine & vbNewLine & "Do you want all of the " & uLevel & "s in this ATF Spec to take on this "
    askMsg4 = vbNewLine & vbNewLine & "Yes or No?"

    Select Case VarType(OldValue)

        Case vbInteger, vbLong, vbDecimal
            askMsg1 = "You changed the " & fLabel & " on a LPAF" & vbNewLine & "From: " & OldValue & " to " & NewValue & "."
            askMsg3 = vbNewLine & "NEW " & fLabel & " of: " & NewValue & "?"
            If uLevel = "Unit" Then SQLz = "UPDATE Units SET " & TargetField & "=" & NewValue & " WHERE " & TargetField & "=" & OldValue & " AND ATFUnitSpecID=" & SpecId & ";"
            If uLevel = "Group" Then SQLz = "UPDATE SCHGRP SET " & TargetField & "=" & NewValue & " WHERE " & TargetField & "=" & OldValue & " AND ATFGroupSpecID=" & SpecId & ";"

        Case vbCurrency
            askMsg1 = "You changed the " & fLabel & " on a LPAF" & vbNewLine & "From: " & Format(OldValue, "$#,##0.00") & " to " & Format(NewValue, "$#,##0.00") & "."
            askMsg3 = vbNewLine & "NEW " & fLabel & " of: " & Format(NewValue, "$#,##0.00") & "?"
            If uLevel = "Unit" Then SQLz = "UPDATE Units SET " & TargetField & "=" & NewValue & " WHERE " & TargetField & "=" & CCur(OldValue) & " AND ATFUnitSpecID=" & SpecId & ";"
            If uLevel = "Group" Then SQLz = "UPDATE SCHGRP SET " & TargetField & "=" & NewValue & " WHERE " & TargetField & "=" & CCur(OldValue) & " AND ATFGroupSpecID=" & SpecId & ";"

        Case vbString
            OldValue = Replace(OldValue, "'", "''")
            NewValue = Replace(NewValue, "'", "''")
            askMsg1 = "You changed the " & fLabel & " on a LPAF" & vbNewLine & "From: '" & OldValue & "' to '" & NewValue & "'."
            askMsg3 = vbNewLine & "NEW " & fLabel & " of: '" & NewValue & "'?"
            If uLevel = "Unit" Then SQLz = "UPDATE Units SET " & TargetField & "='" & NewValue & "' WHERE " & TargetField & "='" & OldValue & "' AND ATFUnitSpecID=" & SpecId & ";"
            If uLevel = "Group" Then SQLz = "UPDATE SCHGRP SET " & TargetField & "='" & NewValue & "' WHERE " & TargetField & "='" & OldValue & "' AND ATFGroupSpecID=" & SpecId & ";"

        Case Else
            ChangeLPAFItems = 2
            Exit Function
    End Select

    Committ = MsgBox(askMsg1 & askMsg2 & askMsg3 & askMsg4, vbQuestion + vbYesNo + vbCritical, "Change " & TargetField & "?")

    If Committ = vbNo Then
        ChangeLPAFItems = False
        Exit Function
    End If

    On Error GoTo FixMe
    DoCmd.SetWarnings True
    DoCmd.RunSQL SQLz
    DoCmd.SetWarnings True

    ChangeLPAFItems = 1
    Exit Function

FixMe:
    ChangeLPAFItems = 2
End Function
