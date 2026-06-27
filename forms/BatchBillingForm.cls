Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public SQLz As String
Public SQLy As String
Public SQLx As String
Public BillIdx As String
Public BillFromx As Date
Public BillToX As Date
Public preSQL, SchSort As String
Dim Committ As String


Private Sub ApplyDateRange_Click()
    SQLz = "SELECT S.ClientID "
    SQLz = SQLz & ",B.BillTypeID, S.ClientShortName AS Client, B.BillType"
    SQLz = SQLz & ",Count(B.BillId) AS Bills, Min(B.BillFrom) AS FirstBill, Max(B.BillFrom) AS LastBill "
    SQLz = SQLz & " FROM Bills AS B "
    SQLz = SQLz & " INNER JOIN (Select SchId,ClientID, ClientShortName "
    SQLz = SQLz & "             FROM vw_SixKeys Group By ClientID,ClientShortName,SchId) AS S ON B.BillSchRef = S.SchID "
    SQLz = SQLz & " GROUP BY S.ClientID, B.BillTypeID, S.ClientShortName, B.BillType, B.BillType "
    SQLz = SQLz & " HAVING Max(BillFrom) >= #" & Me.Billfrom & "# "
    SQLz = SQLz & " AND Min(BillFrom) <= #" & Me.Billto & "# "
    If BillTypeID.Column(0) > 0 Then SQLz = SQLz & " AND BillTypeID=" & Me.BillTypeID.Column(0) & " "
    SQLz = SQLz & " ORDER BY S.ClientShortName, B.BillType;"
    Me.ClientList.RowSource = SQLz
    Me.AssigneeList.RowSource = ""
    Me.BillList.RowSource = ""
    SQLz = "SELECT BillID, BillBankID, BillSchRef, GroupId, BillType, BillTypeID, BillNo, Billdt, Billfrom"
    SQLz = SQLz & ", Billto, BillDueDt, BillSent, BillLink, DueonRec, SchNo, SGrpNO, AssigneeBankID"
    SQLz = SQLz & ", iif(nz(SchBnkClNO,'')='', '',SchBnkClNO & '-') & SgrpBnkNo as BContract "
    SQLz = SQLz & " FROM (((Bills"
    SQLz = SQLz & " INNER JOIN Schedule ON Bills.BillSchRef = Schedule.SchID)"
    SQLz = SQLz & " INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID)"
    SQLz = SQLz & " LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID)"
    SQLz = SQLz & " WHERE BillId IN(0) "
    SQLz = SQLz & " ORDER BY Billdt DESC"
    Me.BatchBillingMassModSub.Form.RecordSource = SQLz
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From Bills_Batch_WorkTable")
    DoCmd.SetWarnings True
End Sub

Private Sub Billfrom_AfterUpdate()
  'Call ResetForm_Click
   Me.Billto = DateSerial(Year(DateAdd("M", 1, Me.Billfrom)), Month(DateAdd("M", 1, Me.Billfrom)) + 1, 1) - 1
  'DateAdd("M", 1, Me.Billfrom)
   Me.Billfrom2 = DateAdd("m", 1, Me.Billfrom)
   Me.BillTo2 = DateAdd("m", 2, Me.Billfrom) - 1
   Me.BillType2 = Me.BillTypeID
End Sub

Private Sub BillTypeID_AfterUpdate()
  Me.BillType = Me.BillTypeID.Column(1)
'Call ResetForm_Click
End Sub

Private Sub CloseMe_Click()
  DoCmd.close acForm, (Me.Form.Name)
End Sub

Public Sub Form_Load()
    Me.AssigneeList.RowSource = ""
    Me.BillList.RowSource = ""
    Me.ClientList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.MissingBillsList.RowSource = ""
    Me.Billfrom2 = ""
    Me.BillTo2 = ""
    Me.BillType2 = ""
    Me.BillTypeID = 0
    Me.GetAll.Visible = False
    Me.GetSelectedBills.Visible = False
    Me.MissingBillsList_Label.caption = "'Stale' or 'Missing' Bills"
End Sub

Public Function MAKE_GroupBillQuery(Optional SQLx As String, Optional NoQuery As Boolean)
    Dim RST As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim i As Integer
    On Error Resume Next
    If SQLx = "" Then
        Set RST = CurrentDb.OpenRecordset("SELECT * FROM Bills_Batch_WorkTable", dbOpenSnapshot)
    Else
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("Delete From Bills_Batch_WorkTable")
        DoCmd.SetWarnings True
        Set RST = CurrentDb.OpenRecordset(SQLx)
    End If
    RST.MoveLast
    RST.MoveFirst
    BillIdx = ""
    For i = 1 To RST.RecordCount
        BillIdx = BillIdx & RST!BillID
        BillFromx = RST!Billfrom
        BillToX = RST!Billto
        If SQLx <> "" Then DoCmd.RunSQL ("Insert into Bills_Batch_WorkTable (BillID) Select" & RST!BillID & ";")
        If RST.RecordCount > 1 And i < RST.RecordCount Then BillIdx = BillIdx & ","
        RST.MoveNext
    Next i
    RST.close
    Set RST = Nothing
    On Error GoTo 0
    If NoQuery = True Then Exit Function

    If Me.ClientList.Column(3) = "PerDiem" Then
        SQLz = "SELECT *"
        SQLz = SQLz & " From vw_PerDiem_Batch"
        SQLz = SQLz & " WHERE BillID In (" & BillIdx & ");"
        On Error Resume Next
        DoCmd.DeleteObject acQuery, "FAPerDiemBatch"
        On Error GoTo 0
        Set qdf = CurrentDb.CreateQueryDef("FAPerDiemBatch", SQLz)
    End If
    If Me.ClientList.Column(3) = "BankBill" Or Me.ClientList.Column(3) Like "*Rent*" Then
        SQLz = "SELECT vw_BankBill_Batch.*"
        SQLz = SQLz & ", CurrRent AS RentNoTax"
        SQLz = SQLz & ", IIf([UnitTaxExempt]=False,CCur([FSLFee]*[LTaxRate]),0) AS FSLTax"
        SQLz = SQLz & ", [CurrRent]+CCur(Nz([FSLFee],0))+CCur(Nz([TaxAmt],0)) AS UnitTotalRent"
        SQLz = SQLz & ", BillToGroup AS BillToAddress, BillNo AS InvoiceNumber"
        SQLz = SQLz & " From vw_BankBill_Batch"
        SQLz = SQLz & " WHERE BillID In (" & BillIdx & ") "
        SQLz = SQLz & " ORDER BY aSort;"
        On Error Resume Next
        DoCmd.DeleteObject acQuery, "FABankBillGroupBatch"
        On Error GoTo 0
        Set qdf = CurrentDb.CreateQueryDef("FABankBillGroupBatch", SQLz)
    End If
     If Me.ClientList.Column(3) = "PDSurrender" Or Me.ClientList.Column(3) = "PDSurrLoc" Then
        SQLz = "SELECT *"
        SQLz = SQLz & " FROM  vw_PerdiemSurrenderInvoice"
        SQLz = SQLz & " WHERE BillID In (" & BillIdx & ");"
        On Error Resume Next
        DoCmd.DeleteObject acQuery, "PerDiemSurrender_Batch"
        On Error GoTo 0
        Set qdf = CurrentDb.CreateQueryDef("PerDiemSurrender_Batch", SQLz)
    End If
End Function
Private Sub ClientList_Click()
    ' Me.UnitsList.RowSource = preSQL & "WHERE ClientId=" & Me.ClientList & " ORDER BY Client, CDbl(NumInString(Left(Sch,3))), Sch, [Group];"
    Me.MissingBillsList_Label.caption = "'Stale' or 'Missing' Bills"
    SQLz = "Select BankID,Assignee, Count(BankID) as Bills From "
    SQLz = SQLz & " (Select BankID,Assignee FROM Bills as B Inner Join"
    SQLz = SQLz & " (Select ClientID,SchId,Schedule, Assignee,BankID From vw_FiveKeys Group By ClientID,SchId,Schedule, Assignee,BankID)  as F "
    SQLz = SQLz & " ON F.SchID=B.BillSchRef "
    SQLz = SQLz & " WHERE BillTypeID=" & Me.ClientList.Column(1) & " "
    SQLz = SQLz & " AND ClientID=" & Me.ClientList.Column(0) & " "
    SQLz = SQLz & " AND BillFrom >= #" & Me.Billfrom & "# "
    SQLz = SQLz & " AND Billto <= #" & Me.Billto & "# "
    SQLz = SQLz & " ) AS B Group BY BankID ,Assignee ORDER BY Assignee "
    Me.AssigneeList.RowSource = SQLz

    SQLz = "Select SchID, Schedule, Count(SchID) as Bills "
    SQLz = SQLz & " FROM Bills as B Inner Join"
    SQLz = SQLz & " (Select ClientID,SchId,Schedule, Assignee,BankID From vw_FiveKeys Group By ClientID,SchId,Schedule, Assignee,BankID)  as F "
    SQLz = SQLz & " ON F.SchID=B.BillSchRef "
    SQLz = SQLz & " WHERE BillTypeID=" & Me.ClientList.Column(1) & " "
    SQLz = SQLz & " AND ClientID=" & Me.ClientList.Column(0) & " "
    SQLz = SQLz & " AND BillFrom >= #" & Me.Billfrom & "# "
    SQLz = SQLz & " AND Billto <= #" & Me.Billto & "# "
    SQLz = SQLz & " GROUP BY SchID, Schedule"
    SQLz = SQLz & " ORDER BY CDbl(NumInString(Left(Schedule,3))) "
    Me.ScheduleList.RowSource = SQLz

    SQLz = "Select BillID,SchID, BillType,Schedule "
    SQLz = SQLz & " ,IIf(NOT b.GroupID IS NULL,(Select SGrpNO From SchGrp WHERE sgrpId=b.GroupID),'') as IsGrp"
    SQLz = SQLz & " ,Assignee ,BillFrom, Billto, BankID, ClientID, BillTypeID"
    SQLz = SQLz & " FROM Bills as B Inner Join"
    SQLz = SQLz & " (Select ClientID,SchId,Schedule, Assignee,BankID From vw_FiveKeys Group By ClientID,SchId,Schedule, Assignee,BankID)  as F "
    SQLz = SQLz & " ON F.SchID=B.BillSchRef "
    SQLz = SQLz & " WHERE BillTypeID=" & Me.ClientList.Column(1) & " "
    SQLz = SQLz & " AND ClientID=" & Me.ClientList.Column(0) & " "
    SQLz = SQLz & " AND BillFrom >= #" & Me.Billfrom & "# "
    SQLz = SQLz & " AND Billto <= #" & Me.Billto & "# "
    SQLz = SQLz & " ORDER BY Assignee, CDbl(NumInString(Left(Schedule,3))), BillFrom "
    Me.BillList.RowSource = SQLz

    Call MAKE_GroupBillQuery(SQLz, True)
    SQLz = "SELECT BillID, BillBankID, BillSchRef, GroupId, BillType, BillTypeID, BillNo, Billdt, Billfrom"
    SQLz = SQLz & ", Billto, BillDueDt, BillSent, BillLink, DueonRec, SchNo, SGrpNO, AssigneeBankID"
    SQLz = SQLz & ", iif(nz(SchBnkClNO,'')='', '',SchBnkClNO & '-') & SgrpBnkNo as BContract "
    SQLz = SQLz & " FROM (((Bills"
    SQLz = SQLz & " INNER JOIN Schedule ON Bills.BillSchRef = Schedule.SchID)"
    SQLz = SQLz & " INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID)"
    SQLz = SQLz & " LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID)"
    SQLz = SQLz & " WHERE BillId IN(0) "
    SQLz = SQLz & " ORDER BY Billdt DESC"
    Me.BatchBillingMassModSub.Form.RecordSource = SQLz

    Me.GetAll.Visible = True
    Me.GetSelectedBills.Visible = True

    SQLz = "SELECT Sch"
    SQLz = SQLz & ", Grp"
    SQLz = SQLz & ", bStatus"
    SQLz = SQLz & ", Age"
    SQLz = SQLz & ", Assignee as Bank"
    SQLz = SQLz & ", BLCD"
    SQLz = SQLz & ", ActiveUnits as Actv"
    SQLz = SQLz & ", LastBill"
    SQLz = SQLz & ", BillType as [Bill Type]"
    SQLz = SQLz & ", FirstPDDate as [1st PD]"
    SQLz = SQLz & ", LastOffLease as Last_EOL"
    SQLz = SQLz & ", clientID"
    SQLz = SQLz & ", SchID"
    SQLz = SQLz & ", groupID"
    SQLz = SQLz & ", uGroupID"
    SQLz = SQLz & ", BillTypeID "
    SQLz = SQLz & " FROM vw_BillsStaleMissing_Batch"
    SQLz = SQLz & " WHERE clientID=" & Me.ClientList.Column(0) & " "
    SQLz = SQLz & " ORDER BY CDbl(NumInString(Left(Sch,3))), GRP "
    Me.MissingBillsList.RowSource = SQLz
    Me.MissingBillsList_Label.caption = Me.ClientList.Column(2) & " 'Stale' or 'Missing' Bills"
End Sub
Private Sub AssigneeList_Click()
    SQLz = "Select BillID,SchID, BillType,Schedule "
    SQLz = SQLz & " ,IIf(NOT b.GroupID IS NULL,(Select SGrpNO From SchGrp WHERE sgrpId=b.GroupID),'') as IsGrp"
    SQLz = SQLz & " ,Assignee ,BillFrom, Billto, BankID, ClientID, BillTypeID"
    SQLz = SQLz & " FROM Bills as B Inner Join"
    SQLz = SQLz & " (Select ClientID,SchId,Schedule, Assignee,BankID From vw_FiveKeys Group By ClientID,SchId,Schedule, Assignee,BankID)  as F "
    SQLz = SQLz & " ON F.SchID=B.BillSchRef "
    SQLz = SQLz & " WHERE BillTypeID=" & Me.ClientList.Column(1) & " "
    SQLz = SQLz & " AND ClientID=" & Me.ClientList.Column(0) & " "
    SQLz = SQLz & " AND BillFrom >= #" & Me.Billfrom & "# "
    SQLz = SQLz & " AND Billto <= #" & Me.Billto & "# "
    SQLz = SQLz & " AND BankID = " & Me.AssigneeList.Column(0) & " "
    SQLz = SQLz & " ORDER BY Assignee, CDbl(NumInString(Left(Schedule,3))), BillFrom "
    Me.BillList.RowSource = SQLz

    Call MAKE_GroupBillQuery(SQLz, False)

    SQLz = "SELECT BillID, BillBankID, BillSchRef, GroupId, BillType, BillTypeID, BillNo, Billdt, Billfrom"
    SQLz = SQLz & ", Billto, BillDueDt, BillSent, BillLink, DueonRec, SchNo, SGrpNO, AssigneeBankID"
    SQLz = SQLz & ", iif(nz(SchBnkClNO,'')='', '',SchBnkClNO & '-') & SgrpBnkNo as BContract "
    SQLz = SQLz & " FROM (((Bills"
    SQLz = SQLz & " INNER JOIN Schedule ON Bills.BillSchRef = Schedule.SchID)"
    SQLz = SQLz & " INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID)"
    SQLz = SQLz & " LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID)"
    SQLz = SQLz & " WHERE BillId IN(" & BillIdx & ") "
    SQLz = SQLz & " ORDER BY Billdt DESC"
    Me.BatchBillingMassModSub.Form.RecordSource = SQLz

    DoCmd.SetWarnings True
    Me.GetAll.Visible = True
    Me.GetSelectedBills.Visible = True
    Me.PickModeToggle = True
    Call PickModeToggle_Click
  ' Call GetAll_Click
End Sub

Private Sub MakeMissingBill_Click()
    If Me.MissingBillsList.Column(0) & "test" = "test" Then Exit Sub
    Committ = MsgBox("You're about to Make a '" & Me.BillType2.Column(1) & "' " & vbNewLine & _
        " for Client: " & Me.ClientList.Column(2) & vbNewLine & _
        " within Schedule: " & Me.MissingBillsList.Column(0) & vbNewLine & _
        " using Bill From: " & Me.Billfrom2 & vbNewLine & _
        "            Bill To  : " & Me.BillTo2 & vbNewLine & vbNewLine & "Are you Sure?", _
        vbQuestion + vbYesNo + vbCritical, "Stop looking at me!")

    If Committ = vbNo Then Exit Sub
    Dim SchIDx As Long
    SchIDx = Me.MissingBillsList.Column(12)
    SQLz = "INSERT INTO Bills (BillSchRef,[GroupId],BillType,BillTypeID,BillNo,Billdt,Billfrom,Billto,BillDueDt,BillAdddt) "
    SQLz = SQLz & "SELECT " & SchIDx & " AS SchID"
    SQLz = SQLz & ", " & Me.MissingBillsList.Column(13) & " as GrpID"
    SQLz = SQLz & ", '" & Me.BillType2.Column(1) & "' as bType"
    SQLz = SQLz & ", " & Me.BillType2.Column(0) & " as BTypeID"
    SQLz = SQLz & ", 'Need Inv Number' as BillNo"
    SQLz = SQLz & ", #" & DatePart("m", Me.Billfrom2) & "/1/" & DatePart("yyyy", Me.Billfrom2) & "#"
    SQLz = SQLz & ", #" & Me.Billfrom2 & "#"
    SQLz = SQLz & ", #" & Me.BillTo2 & "#"
    SQLz = SQLz & ", #" & DateSerial(Year(DateAdd("M", 1, Me.Billfrom2)), Month(DateAdd("M", 1, Me.Billfrom)) + 1, 1) - 1 & "#"
    SQLz = SQLz & ", #" & Date & "#;"
    DoCmd.SetWarnings True
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    SQLz = "SELECT BillID, BillBankID, BillSchRef, GroupId, BillType, BillTypeID, BillNo, Billdt, Billfrom"
    SQLz = SQLz & ", Billto, BillDueDt, BillSent, BillLink, DueonRec, SchNo, SGrpNO, AssigneeBankID"
    SQLz = SQLz & ", iif(nz(SchBnkClNO,'')='', '',SchBnkClNO & '-') & SgrpBnkNo as BContract "
    SQLz = SQLz & " FROM (((Bills"
    SQLz = SQLz & " INNER JOIN Schedule ON Bills.BillSchRef = Schedule.SchID)"
    SQLz = SQLz & " INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID)"
    SQLz = SQLz & " LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID)"
    SQLz = SQLz & " WHERE BillId IN(" & DMax("BillID", "Bills", "BillSchRef=" & SchIDx & "") & ") "
    SQLz = SQLz & " ORDER BY Billdt DESC"
    Me.BatchBillingMassModSub.Form.RecordSource = SQLz

    Me.MissingBillsList.Requery

End Sub

Private Sub ScheduleList_Click()
    SQLz = "Select BillID,SchID, BillType,Schedule "
    SQLz = SQLz & " ,IIf(NOT b.GroupID IS NULL,(Select SGrpNO From SchGrp WHERE sgrpId=b.GroupID),'') as IsGrp"
    SQLz = SQLz & " ,Assignee ,BillFrom, Billto, BankID, ClientID, BillTypeID"
    SQLz = SQLz & " FROM Bills as B Inner Join"
    SQLz = SQLz & " (Select ClientID,SchId,Schedule, Assignee,BankID From vw_FiveKeys Group By ClientID,SchId,Schedule, Assignee,BankID)  as F "
    SQLz = SQLz & " ON F.SchID=B.BillSchRef "
    SQLz = SQLz & " WHERE BillTypeID=" & Me.ClientList.Column(1) & " "
    SQLz = SQLz & " AND SchID=" & Me.ScheduleList.Column(0) & " "
    SQLz = SQLz & " AND BillFrom >= #" & Me.Billfrom & "# "
    SQLz = SQLz & " AND Billto <= #" & Me.Billto & "# "
    SQLz = SQLz & " ORDER BY Assignee, CDbl(NumInString(Left(Schedule,3))), BillFrom "
    Me.BillList.RowSource = SQLz
    Call MAKE_GroupBillQuery(SQLz, False)

    SQLz = "SELECT BillID, BillBankID, BillSchRef, GroupId, BillType, BillTypeID, BillNo, Billdt, Billfrom"
    SQLz = SQLz & ", Billto, BillDueDt, BillSent, BillLink, DueonRec, SchNo, SGrpNO, AssigneeBankID"
    SQLz = SQLz & ", iif(nz(SchBnkClNO,'')='', '',SchBnkClNO & '-') & SgrpBnkNo as BContract "
    SQLz = SQLz & " FROM (((Bills"
    SQLz = SQLz & " INNER JOIN Schedule ON Bills.BillSchRef = Schedule.SchID)"
    SQLz = SQLz & " INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID)"
    SQLz = SQLz & " LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID)"
    SQLz = SQLz & " WHERE BillId IN(" & BillIdx & ") "
    SQLz = SQLz & " ORDER BY Billdt DESC"
    Me.BatchBillingMassModSub.Form.RecordSource = SQLz

    DoCmd.SetWarnings True
    Me.GetAll.Visible = True
    Me.GetSelectedBills.Visible = True
    Me.PickModeToggle = True
    Call PickModeToggle_Click

End Sub
Public Sub GetAll_Click()
    Dim i As Integer
    For i = 0 To BillList.ListCount - 1
        BillList.Selected(i) = True
    Next i
End Sub

Public Sub GetSelectedBills_Click()
  Call MAKE_GroupBillQuery("", False)
    SQLz = "SELECT BillID, BillBankID, BillSchRef, GroupId, BillType, BillTypeID, BillNo, Billdt, Billfrom"
    SQLz = SQLz & ", Billto, BillDueDt, BillSent, BillLink, DueonRec, SchNo, SGrpNO, AssigneeBankID"
    SQLz = SQLz & ", iif(nz(SchBnkClNO,'')='', '',SchBnkClNO & '-') & SgrpBnkNo as BContract "
    SQLz = SQLz & " FROM (((Bills"
    SQLz = SQLz & " INNER JOIN Schedule ON Bills.BillSchRef = Schedule.SchID)"
    SQLz = SQLz & " INNER JOIN MstrLease ON Schedule.SchMLRef = MstrLease.MLID)"
    SQLz = SQLz & " LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID)"
    SQLz = SQLz & " WHERE BillId IN(" & BillIdx & ") "
    SQLz = SQLz & " ORDER BY CDbl(NumInString(Left(SchNo,3))), Billdt DESC"
    Me.BatchBillingMassModSub.Form.RecordSource = SQLz
  End Sub

Private Sub MakeBatchInvoice_Click()
    If Me.ClientList.Column(3) = "PDSurrender" Or Me.ClientList.Column(3) = "PDSurrLoc" Then
        DoCmd.OpenReport "PerDiemSurrender_Batch", acViewPreview
    Else
        DoCmd.OpenReport "FABankBillGroup_Batch", acViewPreview
    End If
End Sub

Private Sub PickModeToggle_Click()
    Dim i As Integer
    For i = Me.BillList.ListCount - 1 To 0 Step -1
        Me.BillList.Selected(i) = 0
    Next i
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From Bills_Batch_WorkTable")
    DoCmd.SetWarnings True
    If Me.PickModeToggle = True Then
        Me.PickModeToggle.caption = ChrW(9660) & "Pick On"
    Else
        Me.PickModeToggle.caption = "Pick Off"
    End If
    Call GetBillCounts
End Sub
Private Sub GetBillCounts()
    Dim Billx, BillZ As String
    If ([BillList].[ListCount]) > 1 Then Billx = "Bill " & ChrW(9658) Else Billx = "Bills " & ChrW(9658)
    If [BillList].[ItemsSelected].[Count] = 1 Then BillZ = "Bills " & ChrW(9658) Else BillZ = "Bills " & ChrW(9658)
    Me.GetAll.caption = ChrW(9660) & "All " & ([BillList].ListCount) & ""    '& BillX
    If ([BillList].[ListCount]) > 0 Then Me.GetAll.Enabled = True Else Me.GetAll.Enabled = False
    Me.GetSelectedBills.caption = ChrW(9660) & "Selected (" & [BillList].[ItemsSelected].[Count] & ") " & BillZ
    If [BillList].[ItemsSelected].[Count] > 0 Then Me.GetSelectedBills.Enabled = True Else Me.GetSelectedBills.Enabled = False
End Sub

Private Sub ResetForm_Click()
   Call Me.Form_Load
   Me.Form.Refresh
End Sub

Private Sub BillList_Click()
       If Me.PickModeToggle = True Then
        DoCmd.SetWarnings False
        If Coalesce(Me.BillList.Column(0), 0) > 0 Then DoCmd.RunSQL ("Insert into Bills_Batch_WorkTable(BillID) SELECT " & Me.BillList.Column(0) & ";")
        DoCmd.SetWarnings True
        Call ReselectPickedBills
    End If
        Call GetBillCounts
  '  Me.GroupList = DLookup("SGrpId", "vw_SixKeys", "UnitID=" & Me.BillList.Column(0) & "")
End Sub
Public Sub ReselectPickedBills()
    Dim i, x As Integer
    Dim pUnits As String
    Dim RST As DAO.Recordset
    SQLz = "Select * From Bills_Batch_WorkTable Order by BillID Desc;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        For x = 1 To RST.RecordCount
            For i = Me.BillList.ListCount - 1 To 0 Step -1
                If Me.BillList.ItemData(i) = CStr(RST!BillID) Then Me.BillList.Selected(i) = 1
            Next i
            RST.MoveNext
        Next x
    End If
    RST.close
    Set RST = Nothing

End Sub
Public Sub ExcelReport_Click()
    Dim XLS As String
    Dim BillITypeX, Header As String

    BillTypeX = Me.ClientList.Column(3)

    If BillTypeX Like "*Bill*" Or BillTypeX Like "*Rent*" Then
        SQLz = "Select "                                                      'Col-A
        SQLz = SQLz & " BillNo As [Invoice Number]"                           'Col-B
        SQLz = SQLz & ",BankContract as [Assignee Contract #]"
        SQLz = SQLz & ",BankName as [Remit To]"                               'Col-D
        SQLz = SQLz & ",MLNO as [Lease:]"
        SQLz = SQLz & ",SchNo as [Sch Num]"
        SQLz = SQLz & ",SGrpNO AS [Grp]"                                      'Col=G
        SQLz = SQLz & ",Billdt as [Invoice Date]"                             'Col-H
        SQLz = SQLz & ",BillDueDt as [Due Date]"
        SQLz = SQLz & ",Billfrom AS [Bill From] "
        SQLz = SQLz & ",Billto As [Bill To]  "
        SQLz = SQLz & ",LocationName as Location"                             'Col-L
        SQLz = SQLz & ",LState "
        SQLz = SQLz & ",UnitUnitNum as Unit"
        SQLz = SQLz & ",UnitVIN as VIN"
        SQLz = SQLz & ",[Desc] as [Description]"
        SQLz = SQLz & ",UnitStatus as Status"
        SQLz = SQLz & ",CurrRent as Rent"
        SQLz = SQLz & ",iif(UnitTaxExempt=True,""N"",""Y"") as [Taxed?]"
        SQLz = SQLz & ",LTaxRate as [Tax Rate]"
        SQLz = SQLz & ",TaxAmt as Tax"
        SQLz = SQLz & ",FSLFee as [Serv Fee]"
        SQLz = SQLz & ",LineTot as [Total Payment] "
        SQLz = SQLz & ",BillNo"
        SQLz = SQLz & " From vw_BankBill_Batch "
        SQLz = SQLz & " WHERE BillID IN (" & BillIdx & ") "
        SQLz = SQLz & " ORDER BY "
        SQLz = SQLz & " CDbl(NumInString(Left(SchNo,3))) desc "
        SQLz = SQLz & ", SchNo"
        SQLz = SQLz & ", SGrpNO"
        Header = Me.ClientList.Column(2) & " Bills From " & Me.Billfrom & " to " & Me.Billto & ""
        XLS = "|N,TC|P,BC|Q,$E|S,P3|T,$E|U,$E|V,$E|V,W13|"
        XLS = XLS & "E,TC|E,W10|F,TC|F,W06|G,W11|V,W12|A,DD|A,Wrap|"
    End If
    If BillTypeX Like "*PerDiem*" Then
        SQLz = " SELECT MLNo"                       'A
        SQLz = SQLz & ",SchNo"
        SQLz = SQLz & ",SGrpNO As GrpNo"
        SQLz = SQLz & ",'FA' as Bank"               'D
        SQLz = SQLz & ",BLCD"
        SQLz = SQLz & ",UnitId As AssetID"
        SQLz = SQLz & ",UnitUnitNum as UnitNum"
        SQLz = SQLz & ",UnitVIN"                    'H
        SQLz = SQLz & ",UnitStatus as Sts"
        SQLz = SQLz & ",UnitDescShortVer as Descp"  'J
        SQLz = SQLz & ",LocationName as Location"
        SQLz = SQLz & ",UnitPDStartDate as UnitPDStart"
        SQLz = SQLz & ",PDBegin As BillPDStart"     'M
        SQLz = SQLz & ",BilledPDEndDate As BillPDEnd"
        SQLz = SQLz & ",UnitTaxExempt as Exempt"
        SQLz = SQLz & ",LTaxRate as TaxRate"
        SQLz = SQLz & ",PDRent" ' P
        SQLz = SQLz & ",PDDays"
        SQLz = SQLz & ",BasePerDiem   as BillPDRent"
        SQLz = SQLz & ",ServiceFee  as ServFee"
        SQLz = SQLz & ",Taxes"
        SQLz = SQLz & ",LineTotal"
        SQLz = SQLz & ",BillNo"
        SQLz = SQLz & " FROM vw_PerDiemInvoice WHERE BillID IN (" & BillIdx & ") Order BY SchNo ;"
        Header = Me.ClientList.Column(2) & " Perdiems From " & Me.Billfrom & " to " & Me.Billto & ""
        XLS = "|I,BC|O,BC|P,P3|Q,$E|R,PC|S,$E|T,$E|U,$E|V,$E|"
    End If
    If BillTypeX Like "*Surrender*" Then
        SQLz = " SELECT MLNo"                       'A
        SQLz = SQLz & ",SchNo"
        SQLz = SQLz & ",SGrpNO"
        SQLz = SQLz & ",'FA' as Bank"               'D
        SQLz = SQLz & ",BLCD"
        SQLz = SQLz & ",UnitId As AssetID"
        SQLz = SQLz & ",UnitUnitNum as UnitNum"
        SQLz = SQLz & ",UnitVIN"                    'H
        SQLz = SQLz & ",UnitStatus as Sts"
        SQLz = SQLz & ",UnitDescShortVer as Descp"  'J
        SQLz = SQLz & ",LocationName as Location"
        SQLz = SQLz & ",UnitPDStartDate as UnitPDStart"
        SQLz = SQLz & ",PDBegin As BillPDStart"     'M
        SQLz = SQLz & ",PDEnd As BillPDEnd"
        SQLz = SQLz & ",UnitTaxExempt as Exempt"
        SQLz = SQLz & ",LTaxRate as TaxRate"
        SQLz = SQLz & ",PDRent" ' P
        SQLz = SQLz & ",PDDays"
        SQLz = SQLz & ",BasePerDiem   as BillPDRent"
        SQLz = SQLz & ",ServiceFee  as ServFee"
        SQLz = SQLz & ",Taxes"
        SQLz = SQLz & ",LineTotal"
        SQLz = SQLz & ",OffLeaseActual"
        SQLz = SQLz & ",InspectionPassed "
        SQLz = SQLz & "FROM vw_PerdiemSurrenderInvoice WHERE BillID IN (" & BillIdx & ") Order BY SchNo ;"
        Header = Me.ClientList.Column(2) & " Perdiems From " & Me.Billfrom & " to " & Me.Billto & ""
        XLS = "|I,BC|O,BC|P,P3|Q,$E|R,PC|S,$E|T,$E|U,$E|V,$E|"
    End If
   Call MessageUser("Open", "Getting Data", "One Moment")
    SendReport2Excel SQLz, "Invoices", Header, True, XLS, 0, True
     Call MessageUser("Close")
End Sub
Public Sub CloneBills_Click()
   SQLz = "Select BillID,BillFrom, Billto"
   SQLz = SQLz & " FROM Bills as B Inner Join"
   SQLz = SQLz & " ( Select ClientID,SchId,Schedule, Assignee,BankID From vw_FiveKeys "
   SQLz = SQLz & "   Group By ClientID,SchId,Schedule, Assignee,BankID)  as F "
   SQLz = SQLz & " ON F.SchID=B.BillSchRef "
   SQLz = SQLz & " WHERE BillTypeID=" & Me.ClientList.Column(1) & " "
   SQLz = SQLz & " AND ClientID=" & Me.ClientList.Column(0) & " "
   SQLz = SQLz & " AND BillFrom >= #" & Me.Billfrom & "# "
   SQLz = SQLz & " AND Billto <= #" & Me.Billto & "# "
   SQLz = SQLz & " ORDER BY BillID "
  Call MAKE_GroupBillQuery(SQLz)
  Call ClonemBills(BillIdx, _
                    DatePart("m", BillFromx) + 1 & "/1/" & DatePart("yyyy", BillFromx), _
                    DateAdd("m", 1, BillFromx), _
                    DateAdd("m", 1, BillToX), _
                    DateAdd("m", 1, BillFromx))
End Sub
Public Function ClonemBills(BillIdx As String, nBilldate As Date, nBillFrom As Date, nBillto As Date, nDueDate As Date)

    Committ = MsgBox("You're about to Clone '" & Me.BillTypeID.Column(1) & "s'" & vbNewLine & _
             " for Client: " & Me.ClientList.Column(2) & vbNewLine & _
             " using Bill From: " & nBillFrom & vbNewLine & _
             "            Bill To  : " & nBillto & vbNewLine & vbNewLine & "Are you Sure?", _
              vbQuestion + vbYesNo + vbCritical, "Automation is pure fire!")
    If Committ = vbNo Then Exit Function

    SQLz = "INSERT INTO Bills (BillSchRef,[GroupId],BillType,BillTypeID,BillNo,Billdt,Billfrom,Billto,BillDueDt,BillAdddt)"
    SQLz = SQLz & "SELECT BillSchRef AS SchID, GroupID, BillType, BillTypeID"
    SQLz = SQLz & ", 'Needs Inv No' AS BillNo"
    SQLz = SQLz & ", '" & nBilldate & "' AS nBilldate"
    SQLz = SQLz & ", '" & nBillFrom & "' AS nBillFrom"
    SQLz = SQLz & ", '" & nBillto & "' AS nBillto"
    SQLz = SQLz & ", '" & nDueDate & "' AS nDueDate"
    SQLz = SQLz & ", '" & Date & "' AS nDateAdded "
    SQLz = SQLz & "FROM Bills "
    SQLz = SQLz & "WHERE BillID In (" & BillIdx & ");"
On Error GoTo FixMe
   ' DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    SQLz = "UPDATE BILLS Set BILLS.BillBankID=x.AssigneeBankID "
    SQLz = SQLz & "FROM (SELECT BillID, BillBankID, BillSchRef, Bills.GroupId, BillType, BillTypeID, BillNo"
    SQLz = SQLz & "    , Billdt, Billfrom, NZ(AssigneeBankID,BankID) as AssigneeBankID "
    SQLz = SQLz & "      FROM ((Bills INNER JOIN vw_FiveKeys ON Bills.BillSchRef = vw_FiveKeys.SchID)"
    SQLz = SQLz & "      LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID) WHERE BillAdddt = #" & Date & "# "
    SQLz = SQLz & "       AND BillBankID IS NULL) AS X "
    SQLz = SQLz & "WHERE BILLS.BillID = x.BillID;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Me.Billfrom = nBillFrom
    Me.Billto = nBillto
    Call GetSelectedBills_Click
    Me.ClientList = Me.ClientList.Column(0)
    Call ClientList_Click

Exit Function
FixMe:
    MsgBox Err.Description, vbCritical, "(Error!)" & AtlaasVersion
End Function
