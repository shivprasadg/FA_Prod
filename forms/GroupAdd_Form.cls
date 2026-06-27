Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SchIDx, ClientIDx, MLIDx, GroupIDi, SourceGroupID, NextGrpi As Integer
Dim SQLz, SchNoX, FAIDx As String
Dim Msgx   As String





Public Sub Form_Load()
    If IsNull(Me.OpenArgs) Then
        SchIDx = 927
    Else
        SchIDx = SplitString(Me.OpenArgs, "|", 0)
        SourceGroupID = SplitString(Me.OpenArgs, "|", 1)
    End If
    On Error Resume Next
    NextGrpi = Asc(Right(DMax("SGrpNO", "SchGrp", "SGrpSchID=" & SchIDx), 1)) + 1
    SchNoX = DLookup("SchNo", "Schedule", "SchID=" & SchIDx)
    If NextGrpi = 0 Then NextGrpi = Asc("A")
    Me.SchID = SchIDx
    Me.SourceGrpID = CInt(SourceGroupID)
    'SourceGroupID = DMax("SGrpID", "SchGrp", "SGrpSchID=" & SchIdx)

   ' SQLz = "SELECT SGrpSchID, ATFGroupSpecID, 0 as UnitsInGroup, OrigEquipCost, LeasePymtForGroup, LRFGroup, ExchangeITGroup, ExchangeITType, ExchangeITMonth, LeaseExtOption1Term, " & _
           "LeaseTermPrimary, 3 as LeaseTermNotice, LeasePymtFreq, LeasePaidAdvArr, LeasePymtForGroup, SGrpSalesIndex, SGrpSalesIndDt, SgrpSalesIndRt, SGprSalesIndYr, " & _
           "SgrpSalesSpread, SgrpSalesYield, SGrpBonus, BPChgLO, RebateUnitProjected, SGRPPLB, LeaseType, SgrpCommInd, PerDiemStartsOn, AllowedHoursPer, SGprAllow1, SGrpov1, SGprAllow2, SGrpov2, SGprAllow4, SGrpov4, SGprAllow5, SGrpov5, SGrpFSLPmt, ScheduleGroupPaymentTerm, SGrpRSLTPmt " & _
           "From SchGrp WHERE  sGrpId=" & SourceGroupID & ""
     Me.Form.RecordSource = ""
     DoCmd.SetWarnings False
     DoCmd.RunSQL ("Drop Table GroupFormAdd_WorkTable")
     'DoCmd.RunSQL ("INSERT INTO GroupFormAdd_WorkTable SELECT * FROM Schgrp Where SGrpID=" & SourceGroupID & ";")
     DoCmd.RunSQL ("SELECT * INTO GroupFormAdd_WorkTable FROM SchGrp Where SGrpID=" & SourceGroupID & ";")
     SQLz = "Select * From GroupFormAdd_WorkTable"
     DoCmd.SetWarnings False
    Me.Form.RecordSource = (SQLz)
    Me.SGrpNO = Chr(NextGrpi)
    'Me.SGrpFAID = FAIDx
    Me.FormLabel.caption = "Adding Group " & Chr(NextGrpi) & " to Schedule " & SchNoX

    ClientIDx = DLookup("ClientID", "vw_SixKeys", "SchID=" & SchIDx)
    MLIDx = DLookup("MLID", "vw_SixKeys", "SchID=" & SchIDx)
    SchNoX = DLookup("SchNo", "Schedule", "SCHId=" & SchIDx)
End Sub
Private Sub CloneSchNGroup_Click()
Msgx = ""
 If Me.NewSchNo & "test" = "test" Then Msgx = Msgx & "    New SchNo is missing" & vbNewLine
 If Me.NewGrpNo & "test" = "test" Then Msgx = Msgx & "    New GrpNo is missing" & vbNewLine
     If Msgx <> "" Then
        Msgx = "Please correct the the following:" & vbNewLine & Msgx & vbNewLine
        Msgx = Msgx & "The Schedule and the Group have NOT been cloned..."
        MsgBox Msgx, vbCritical, AtlaasVersion
        Exit Sub
    End If
Call CloneScheduleGroup(Me.SchID, Me.SourceGrpID, Me.NewSchNo, Me.NewGrpNo, 0, 0, 0, 0, True)

End Sub

Private Sub MakeGroup_Click()
    Msgx = ""
    If Me.SGrpNO & "test" = "test" Then Msgx = Msgx & "    Group No is missing" & vbNewLine
    If Me.SGrpFAID & "test" = "test" Then Msgx = Msgx & "    FAID is missing" & vbNewLine
    If Me.UnitsInGroup & "test" = "test" Then Msgx = Msgx & "    Units in Group is missing" & vbNewLine
    If Me.OrigEquipCost & "test" = "test" Then Msgx = Msgx & "    OEC Cost is missing" & vbNewLine
    If Me.LeaseTermPrimary & "test" = "test" Then Msgx = Msgx & "    Lease Term is missing" & vbNewLine
    If Me.LeaseTermNotice & "test" = "test" Then Msgx = Msgx & "    Lease Notice is missing" & vbNewLine
    If Me.LeasePymtFreq & "test" = "test" Then Msgx = Msgx & "    Pymt Freq is missing" & vbNewLine
    If Me.LeasePaidAdvArr & "test" = "test" Then Msgx = Msgx & "    Paid Adv/Arr is missing" & vbNewLine
    If Me.LeasePymtForGroup & "test" = "test" Then Msgx = Msgx & "    Pymt For Group is missing" & vbNewLine
    If Me.SGrpNO <> "A" Then
        If Me.Reason & "test" = "test" Then Msgx = Msgx & "    Reason For Adding Group is missing" & vbNewLine
        If Len(Me.Reason) > 0 And Len(Me.Reason) < 10 Then Msgx = Msgx & "    Reason For Adding Group is too short" & vbNewLine
    End If
    If Msgx <> "" Then
        Msgx = "Please correct the the following:" & vbNewLine & Msgx & vbNewLine
        Msgx = Msgx & "Group '" & Me.SGrpNO & "' has NOT been added in..."
        MsgBox Msgx, vbCritical, AtlaasVersion
        Exit Sub
    End If
    ' NextGrpi = Asc(Nz(Right(DMax("SGrpNO", "SchGrp", "SGrpSchID=" & Me.SchID), 1), "A") + 1)
Deploy_Group:
    Call InsertGroup(Me.SGrpFAID, Me.SGrpNO, SourceGroupID, Me.SchID, Me.Form.Name, True)
    GroupIDi = DMax("sGrpID", "SchGrp", "sGrpSchID=" & Me.SchID)
    If Me.SGrpNO <> "A" Then Call InsertChangeReason(Date, Me.Reason, "Structure", "Added Group", "schgrp", "Added A Group", Me.SGrpNO, SchNoX, 0, 0, GroupIDi, Me.SchID, 0)
    MsgBox "Group '" & Me.SGrpNO & "' as been added into Schedule " & DLookup("SchNo", "Schedule", "SCHId=" & Me.SchID) & vbNewLine & " The form will close...", vbInformation, atlaasverision
    DoCmd.OpenForm "ScheduleControl"
    If IsOpen("Schedule_Form") Then
        Forms!Schedule_Form.Form.Requery
    End If
    DoCmd.close acForm, Me.Form.Name
End Sub


Private Sub SGrpNO_AfterUpdate()

If DCount("sGrpNo", "SchGrp", "sgrpSchID = " & Me.SchID & " AND SGrpNO='" & Me.SGrpNO & "'") > 0 Then
NextGrpi = Asc(Right(DMax("SGrpNO", "SchGrp", "SGrpSchID=" & SchIDx), 1)) + 1
MsgBox "There is a Group " & Me.SGrpNO & " in this Schedule " & vbNewLine & "Altlaas suggest  '" & Chr(NextGrpi) & "' as the the next Group No."
Me.SGrpNO = Chr(NextGrpi)
End If
Me.FormLabel.caption = "Add Group " & Me.SGrpNO & " to Schedule " & DLookup("SchNo", "Schedule", "SCHId=" & SchIDx)
End Sub
