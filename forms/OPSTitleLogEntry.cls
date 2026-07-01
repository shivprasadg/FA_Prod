Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub AddReason_Click()
DoCmd.close acForm, "AddLogReason"
DoCmd.OpenForm "AddLogReason"

End Sub
Private Sub CheckEntryType()
MsgBox "You Must Select A Entry Type, Either IN or OUT", vbInformation, "Rules.. Rule"
Me.EntryReason.SetFocus
End Sub

Private Sub CloseMe_Click()
Call SaveMe_Click
End Sub

Private Sub DeleteLog_Click()
If AskToDoTask("You are about to delete a TITLE Log...") = True Then
SQLz = "Delete * From UnitTitleLog Where UnitTitleLogId =" & Me.UnitTitleLogID
DoCmd.SetWarnings False
DoCmd.RunSQL (SQLz)
DoCmd.SetWarnings True
If IsOpen("UnitsForm") Then Forms!Unitsform.UnitTitleLogForm.Form.Requery
DoCmd.close acForm, Me.Form.Name
Else
Exit Sub
End If
End Sub

Private Sub EntryReason_Click()
Me.EntryTypeID = Me.EntryReason.Column(1)
Me.ExpectedReturn = Me.EntryReason.Column(2)
Me.EntryReasonId = Me.EntryReason.Column(3)
If Not IsNull(Me.TitleState) Then Me.SaveMe.Enabled = True Else Me.SaveMe.Enabled = False

End Sub

Private Sub EntryReasonId_GotFocus()
If IsNull(Me.EntryTypeID) Then Call CheckEntryType
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo FixMe
If IsNull(Me.EntryTypeID) Then
Me.Form.Undo
End If
FixMe:
End Sub

Private Sub Form_Load()
Me.SaveMe.Enabled = False
If FAUserAccessLevel > 15 Then Me.DeleteLog.Visible = True Else Me.DeleteLog.Visible = False
If Me.USERID & "test" = "test" Then Me.USERID = FAUserIDx
End Sub

Private Sub SaveMe_Click()
    Dim Msgx As String
    If Me.EntryReason & "test" = "test" Then Msgx = Msgx & vbNewLine & "You must select a Reason"
    If UnitTitleLogDate & "test" = "test" Then Msgx = Msgx & vbNewLine & "You must select a Log Date"
    If EntryTypeID & "test" = "test" Then Msgx = Msgx & vbNewLine & "You must select Entry Type  IN or OUT"
    If Me.TitleState & "test" = "test" And EntryTypeID <> 41 Then Msgx = Msgx & vbNewLine & "You must select a State"
    If Msgx <> "" Then
        MsgBox "There is a problem with you log entry:" & Msgx & vbNewLine & vbNewLine & "...the record did not save.", vbCritical, "Missing Title Log info"
        Exit Sub
    End If
    If Me.Dirty Then Me.Dirty = False
    If Me.EntryReason.Column(3) = 36 Then
        Dim MatchingUnitID As Long
        MatchingUnitID = DLookup("NewAssetID", "vw_ReLeasedUnits", "AssetID=" & Me.UnitRef)
        If DCount("UnitTitleLogId", "UnitTitleLog", "EntryReasonId=39 and UnitREf=" & MatchingUnitID) > 0 Then Exit Sub
        Dim msg1   As String
        msg1 = msg1 & "<b><Font Size=3>Title Event Copy/Clone</font></b><br/>"
        msg1 = msg1 & "Atlaas Core is about to apply the same title event<br /><br />"
        msg1 = msg1 & "  <b>'Reallocate/Remain with Custodian'</b><br/>"
        msg1 = msg1 & "  To matching AssetID:<b> " & MatchingUnitID & ".</b><br /><br/>Do you agree?"
        ResposeX = Dialog.RichBox(msg1, vbCustom + vbExclamation, AtlaasVersion, , , 0, 1, 0, 0, "Yes", "No")
        If ResposeX = vbBt1 Then ' Yes
            SQLz = "INSERT INTO UnitTitleLog (UnitRef,EntryTypeId,EntryReasonId,EntryReason,TitleNumber,TitleSt,UnitTitleLogDate,ExpectedReturn,UserID) "
            SQLz = SQLz & "Select " & MatchingUnitID & "," & Me.EntryTypeID & "," & EntryReasonId & ",'" & Me.EntryReason & "','" & Me.TitleNumber & "','" & Me.TitleSt & "','" & Me.UnitTitleLogDate & "'," & ExpectedReturn & ",'" & FaUserName & "';"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
        End If
    End If

    Me.Form.Requery
    If IsOpen("UnitsForm") Then Forms!Unitsform.UnitTitleLogForm.Form.Requery
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub TitleNumber_GotFocus()
If IsNull(Me.EntryTypeID) Then Call CheckEntryType
End Sub

Private Sub TitleState_Click()
If Me.TitleState & "TEST" <> "TEST" Then Me.SaveMe.Enabled = True Else Me.SaveMe.Enabled = False
End Sub

Private Sub TitleState_GotFocus()
'If IsNull(Me.EntryTypeId) Then Call CheckEntryType
End Sub
