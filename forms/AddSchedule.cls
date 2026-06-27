Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CloseMe_Click()
    On Error GoTo FixMe
    DoCmd.close acForm, Me.Form.Name
    Exit Sub
FixMe:
    MsgBox "A Schedule was NOT Added", vbCritical, "Un-Done"
    Me.Form.Undo
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub GoHome_Click()
    DoCmd.OpenForm "FAMainMenu"
End Sub

Private Sub MakeSchedule_Click()

    On Error GoTo MakeSchedule_Error
    Dim Msgx   As String

    If Me.MLID & "test" = "test" Then Msgx = Msgx & "    Master Lease is missing" & vbNewLine
    If Me.SchNo & "test" = "test" Then Msgx = Msgx & "    Schedule Num is missing" & vbNewLine
    If Me.SchShortDesc & "test" = "test" Then Msgx = Msgx & "    Schedule Desc is missing" & vbNewLine
    If Me.NumOfUnits & "test" = "test" Then Msgx = Msgx & "    Units In Group is missing" & vbNewLine
    If Me.Reason & "test" = "test" Then Msgx = Msgx & "    Reason For Adding Schedule is missing" & vbNewLine
    If Len(Me.Reason) > 0 And Len(Me.Reason) < 10 Then Msgx = Msgx & "    Reason For Adding Schedule is too short" & vbNewLine
    If Msgx <> "" Then
        Msgx = "Please correct the the following:" & vbNewLine & Msgx & vbNewLine
        Msgx = Msgx & "The schedule has NOT been added in..."
        MsgBox Msgx, vbCritical, AtlaasVersion
        Exit Sub
    End If
    If DCount("Schedule", "vw_SixKeys", "MLID=" & Me.MLID & " AND Schedule='" & Me.SchNo & "'") > 0 Then
        MsgBox "Schedule '" & Me.SchNo & "' already exist within" & vbNewLine & "Master Lease: " & Me.SelecetMasterLease.Column(2) & vbNewLine & vbNewLine & "The schedule has NOT been added in...", vbCritical, AtlaasVersion
        Exit Sub
    End If

Deploy_Schedule:
    Call InsertSchedule(Me.MLID, Me.SourceSchID, Me.SchNo, Me.SchShortDesc)
    Dim SchIdi As Long
    SchIdi = DMax("SchID", "Schedule", "SchFAID ='" & Me.FAID & "'")
    If DCount("ATFID", "Schedule", "SchID =" & Me.SourceSchID) = 0 Then
        Exit Sub
    End If
Deploy_Group:
    Call InsertGroup(Me.FAID, "A", DMax("SchGrpID", "vw_sixKeys", "Schid=" & Me.SourceSchID), SchIdi)

Save_Reasons:
    GroupIDx = DMax("sgrpid", "schgrp", "SGrpFAID='" & Me.FAID & "'")
    Call InsertChangeReason(Date, Me.Reason, "Structure", "Schedule", "Schedule", "Added Schedule", "", Me.SchNo, "-", Me.SchNo, GroupIDx, SchIdi, 0)
    Call InsertChangeReason(Date, Me.Reason, "Structure", "Group", "schGrp", "Added A Group", "A", Me.SchNo, "-", "A", GroupIDx, SchIdi, 0)
    On Error GoTo 0
    Exit Sub
MakeSchedule_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MakeSchedule_Click of VBA Document Form_AddSchedule"
'Insert_Schedule_Profile:
'    Dim ATFIDx As Integer
'    Dim ATFSchSpecIDX As Integer
'    Dim NextGrpi As Integer
'    On Error Resume Next

 '   ATFIDx = DLookup("ATFID", "Schedule", "SchID =" & Me.SourceSchID)
 '   If ATFIDx > 0 Then
 '       ATFSchSpecIDX = DLookup("ATFSchSpecID", "Schedule", "SchID =" & Me.SourceSchID)
 '       Call InsertSchedule_Profile(Me.MLID, ATFIDx, Me.SchNo, Me.FAID)
 '   Else
 '       ATFSchSpecIDX = 0
 '   End If

'Insert_Group_Profile:
'    If ATFSchSpecIDX > 0 Then
'        NextGrpi = Asc(Right(DMax("ATFGroupSpecDesc", "ATFGroupSpec", "ATFSchSpecID=" & ATFSchSpecIDX), 1)) + 1
'        Call InsertGroup_Profile(ATFSchSpecIDX, Me.NumOfUnits, "A")
'        Dim ATFGroupSpecIDx As Integer
 '       ATFGroupSpecIDx = DMax("ATFGroupSpecID", "ATFGroupSpec", "ATFSchSpecID=" & ATFSchSpecIDX)
 '   Else
  '      ATFGroupSpecIDx = 0
  '  End If
End Sub
