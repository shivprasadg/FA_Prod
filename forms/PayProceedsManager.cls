Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz       As String
Dim SQLt       As String
Dim Schx      As String
Dim UnitsOrPartsX As String
Dim UnitsOrPartsVendor As String
Dim RST        As DAO.Recordset
Public Sub ClearBoxes()
    Dim Ccrl   As Control
    For Each Ccrl In Me.Controls
        If Me(Ccrl.Name).Tag = "Medit" Then
            Me(Ccrl.Name).Value = ""
        End If
    Next Ccrl
End Sub

Private Sub ClearClientSearch_Click()
    Me.SearchClient = ""
    Call SearchClient_Change
End Sub
Private Sub ClearWorkTable()
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From SchIDWorkTable;")
    DoCmd.SetWarnings True
End Sub
Private Sub CloseMe_Click()
    DoCmd.OpenForm "FAMainMenu"
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Editdraw_Click()
     If Me.DrawList & "test" = "test" Then Exit Sub
    DoCmd.close acForm, "DrawDownAddaDraw"
    Call DoCmd.OpenForm("DrawDownAddaDraw", , , , , , "Edit>" & Me.DrawList)
End Sub

Private Sub ExcelUnitPartReport_Click()
If Schx & "test" = "test" Then Exit Sub
    Call MakeDrawDownUnitReport("SchID", CLng(Schx), "Unit Draw-Down Report - for Client:" & DLookup("ClientShortName", "vw_SixKeys", "SchId=" & Schx) & "Sch: " & DLookup("Schedule", "vw_SixKeys", "SchId=" & Schx))
End Sub

Private Sub Form_Load()
    Me.SchNo = "Please select a client..."
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
           "From vw_SixKeys WHERE MLOrig='FA' " & _
           "GROUP BY clientID, ClientShortName, CompanyName " & _
           "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz

    Call ClearSearch("All")
    Call ClearBoxes
    Call ClearWorkTable
    Me.LeaseList.RowSource = ""
    Me.ScheduleList.RowSource = ""
    Me.DrawList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.UnitCount.Visible = False
End Sub

Private Sub GetUnitCounts()
    Dim UnitX, UnitZ As String
    If ([Units_List].[ListCount] - 1) = 1 Then UnitX = "Unit " & ChrW(9658) Else UnitX = "Units " & ChrW(9658)
    If [Units_List].[ItemsSelected].[Count] = 1 Then UnitZ = "Unit " & ChrW(9658) Else UnitZ = "Units " & ChrW(9658)
End Sub
Public Sub ClientList_Click()
    SQLz = "SELECT MLID, MLNo, MLOrig, ClientID " & _
           "From vw_SixKeys " & _
           "WHERE MlOrig='FA' " & _
           "GROUP BY MLID, MLNo, MLOrig, ClientID Having ClientID =" & Me.ClientList & " ORDER BY MLNo;"
    Me.LeaseList.RowSource = SQLz

    Me.ScheduleList.RowSource = ""
    Me.DrawList.RowSource = ""
    Me.Units_List.RowSource = ""
    Me.SchID = Null
    Me.GroupID = Null
    Me.UnitID = Null
    Me.SchNo = Me.ClientList.Column(1)
    Call ClearWorkTable
    'Call GetUnitCounts
End Sub

Private Sub GetSelectedUnits_Click()
    Call ClearWorkTable
    If Me.PickModeToggle = False Then
        Schx = Me.ScheduleList.Column(0)
        DoCmd.SetWarnings False
        xSql = "INSERT INTO SchIDWorkTable(wSchID) Select " & Schx & " as wSchID;"
        DoCmd.RunSQL (xSql)
        DoCmd.SetWarnings True
    Else

        SQLz = "Select SchID From Pick_Schedule_WorkTable Group By SchID Order by SchID Desc;"
        Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
        On Error Resume Next
        RST.MoveLast
        RST.MoveFirst
        On Error GoTo 0
        If RST.RecordCount = 0 Then
            MsgBox "You must 1st Select a Schedule... " & vbNewLine & "did you enable the 'Pick' button?"
            Exit Sub
        End If
        If RST.RecordCount > 0 Then
            Schx = ""
            For x = 1 To RST.RecordCount
                Schx = Schx & RST!SchID & ","
                DoCmd.SetWarnings False
                xSql = "INSERT INTO SchIDWorkTable(wSchID) Select " & RST!SchID & " as wSchID;"
                DoCmd.RunSQL (xSql)
                DoCmd.SetWarnings True
                RST.MoveNext
            Next x
            Schx = Left(Schx, (Len(Schx) - 1))
        End If
        RST.close
        Set RST = Nothing
    End If
    On Error GoTo 0

    SQLz = "Select AssetID"
    SQLz = SQLz & ",UnitNum"
    SQLz = SQLz & ",VIN"
    SQLz = SQLz & ",ItemDesc"
    SQLz = SQLz & ",[Sch|Grp]"
    SQLz = SQLz & ",Vendor"
    SQLz = SQLz & ",IFBank"
    SQLz = SQLz & ",VendorCost"
    SQLz = SQLz & ",uIFAmt"
    SQLz = SQLz & ",iFPts"
    SQLz = SQLz & ",PartCost"
    SQLz = SQLz & ",PartIFAmt"
    SQLz = SQLz & ",IFAmount"
    SQLz = SQLz & ",uIFVar "
    SQLz = SQLz & ",cIFVar "
    SQLz = SQLz & " FROM vw_PayProceeds_IfUnits as PPIF"
    SQLz = SQLz & " WHERE SchID In (" & Schx & ")"
    SQLz = SQLz & " ORDER BY IFBank DESC;"
    Me.Units_List.RowSource = SQLz

    Call getUnitTotals(Schx)

    SQLz = "SELECT Sch"
    SQLz = SQLz & " ,iIF(Units> 0,'Unit','Part') As [Type]"
    SQLz = SQLz & " ,Sum(vUnits + vParts) as vCnt"
    SQLz = SQLz & " ,Sum(PayVendor) as ToVendor"
    SQLz = SQLz & " ,[InterimBank] as iFBank"
    SQLz = SQLz & " ,Sum((iUnits + iParts)) as iCnt"
    SQLz = SQLz & " ,Sum(PPS.InterimAmt) as InterimAmt"
    SQLz = SQLz & " ,Sum((faUnits + faParts)) as faCnt"
    SQLz = SQLz & " ,Sum([FAAmount]) as [FA Cash]"
    SQLz = SQLz & " From vw_PayProceeds_Summary as PPS"
    SQLz = SQLz & " WHERE Schid In (" & Schx & ")"
    SQLz = SQLz & " GROUP BY Sch, iIF(Units> 0,'Unit','Part'), [InterimBank]"
    SQLz = SQLz & " ORDER BY Sch, iIF(Units> 0,'Unit','Part') DESC,[InterimBank];"

    Me.PayOutList.RowSource = SQLz

    SQLz = "Select DrawDownID, DrawNumber AS Num, DrawName, PPCertSent as CertSent "
    SQLz = SQLz & "From DrawDowns "
    SQLz = SQLz & "WHERE DrawDownID "
    SQLz = SQLz & "IN  (Select DrawDownID"
    SQLz = SQLz & "     FROM vw_PayProceeds_Base AS PPB"
    SQLz = SQLz & "     WHERE Schid In (" & Schx & ")"
    SQLz = SQLz & "     Group BY DrawDownID);"
    Me.DrawList.RowSource = SQLz
End Sub
Public Function getUnitTotals(xSch As String)
    Me.UnitIF = DSum("IFAmount", "vw_PayProceeds_IfUnits", "SchID In (" & Schx & ")")

End Function
Public Sub Leaselist_Click()

    SQLz = "SELECT PPIF.SchID"
    SQLz = SQLz & ", PPIF.Sch"
    SQLz = SQLz & ", B.iFBank"
    SQLz = SQLz & ", Sum(PPIF.Unit) AS Units"
    SQLz = SQLz & ", IIf(Max(iType)='Unit',Sum(PayInterim),0) AS iUnits"
    SQLz = SQLz & ", IIf(Min(iType)='Part',Sum(PayInterim),0) AS iParts"
    SQLz = SQLz & " FROM (vw_PayProceeds_UnitParts_Union AS PPIF "
    SQLz = SQLz & " INNER JOIN (SELECT SchID, SchMLRef AS MLID           "
    SQLz = SQLz & " FROM   Schedule)  AS M ON PPIF.SCHID = M.SchID) "
    SQLz = SQLz & " LEFT JOIN (Select BankID, BankShortName as iFBank "
    SQLz = SQLz & " FROM Banks) AS B ON PPIF.InterimBankID = B.BankID "
    SQLz = SQLz & " WHERE (((m.[MLID]) = " & Me.LeaseList & ")) "
    SQLz = SQLz & " GROUP BY PPIF.SchID, PPIF.Sch, B.iFBank "
    SQLz = SQLz & " ORDER BY CDbl(NumInString(Left(Sch,3))) DESC , PPIF.Sch DESC;"

    Me.ScheduleList.RowSource = SQLz

    Me.SchID = Null
    Me.GroupID = Null
    Me.UnitID = Null

    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1)
    Me.DrawList.RowSource = ""
    Me.Units_List.RowSource = ""
    ' Call GetUnitCounts
End Sub
Public Function MAkeReportQuery(QueryType As String)
    Dim SourceQuery As String
    Dim TotProceeds As Currency
    UnitsOrPartsX = ""
    UnitsOrPartsVendor = ""
    SQLt = ""
    If QueryType Like "*Cert*" Then QueryType = "_Cert" Else QueryType = "_Letter"
    SourceQuery = "vw_PayProceeds_Summary" & QueryType
    UnitsOrPartsX = Nz(ConcatRelated("ItemDesc", "vw_PayProceeds_Interim_Header", "iType='Unit' AND InterimAmt > 0 AND SchID in (" & Schx & ")", "iType DESC,ItemDesc", "", "", "), "), "")
    If UnitsOrPartsX = "" Then
        UnitsOrPartsX = Nz(ConcatRelated("ItemDesc", "vw_PayProceeds_Interim_Header", "iType='Part' AND InterimAmt > 0 AND SchID in (" & Schx & ")", "iType DESC,ItemDesc", "", "", "), "), "")
    End If

    UnitsOrPartsVendor = Nz(ConcatRelated("ItemDesc", "vw_PayProceeds_Vendor_Header", "iType='Unit' AND SchID in (" & Schx & ")", "iType DESC,ItemDesc", "", "", "), "), "")
    If UnitsOrPartsVendor = "" Then
        UnitsOrPartsVendor = Nz(ConcatRelated("ItemDesc", "vw_PayProceeds_Vendor_Header", "iType='Part' AND SchID in (" & Schx & ")", "iType DESC,ItemDesc", "", "", "), "), "")
    End If


    TotProceeds = DSum("SchProceeds", "vw_PayProceeds_TotalProceeds", "SCHID IN(" & Schx & ")")

    UnitsOrPartsX = "Schedules:<br/>" & UnitsOrPartsX & ""
    SQLt = "SELECT * "
    SQLt = SQLt & ",'" & UnitsOrPartsX & "'  as [Schs] "
    SQLt = SQLt & ",'" & UnitsOrPartsVendor & "'  as [VendorSchs] "
    SQLt = SQLt & ",'" & TotProceeds & "'  as [TotalProceeds] "
    SQLt = SQLt & " FROM vw_PayProceeds_Summary"
    SQLt = SQLt & " WHERE SchID in (" & Schx & ")"
    If QueryType = "_Cert" Then SQLt = SQLt & " AND InterimAmt > 0 "
   ' SQLt = SQLt & " AND ([iUnits]+[iParts])>0 AND InterimAmt > 0"
    Set qdf = CurrentDb.CreateQueryDef("PayProceeds_data" & QueryType, SQLt)

    SQLt = "SELECT *, '" & UnitsOrPartsX & "'  as [Schs] FROM vw_PayProceeds_UnitParts_Union "
    SQLt = SQLt & " WHERE SchID in (" & Schx & ")"
    If QueryType = "_Cert" Then SQLt = SQLt & " AND InterimAmt > 0 "
    Set qdf = CurrentDb.CreateQueryDef("PayProceeds_data_Units" & QueryType, SQLt)
End Function
Private Sub MakePPLetter_Click()
    DoCmd.DeleteObject acQuery, "PayProceeds_data_Letter"
    DoCmd.DeleteObject acQuery, "PayProceeds_data_Units_Letter"

    Call MAkeReportQuery("Letter")

    DoCmd.close acReport, "PayProceeds_Letter"
    DoCmd.OpenReport "PayProceeds_Letter", acViewReport
    Exit Sub
    Dim myApp  As Word.Application
    Dim myDoc  As Word.Document
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputReport, "PayProceeds_Letter", acFormatPDF, "C:\FAAS\" & FaUserName & "\temp\PayProceeds_Letter.pdf"
    DoCmd.SetWarnings True
    Set myApp = CreateObject("Word.Application")
    Set myDoc = myApp.Documents.Add("C:\FAAS\" & FaUserName & "\temp\PayProceeds_Letter.pdf")
    myApp.Visible = True
End Sub
Private Sub MakePPCert_Click()
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "PayProceeds_data_Cert"
    DoCmd.DeleteObject acQuery, "PayProceeds_data_Units_Cert"
    On Error GoTo 0
    Call MAkeReportQuery("Cert")
    DoCmd.close acReport, "PayProceeds_Cert"
    DoCmd.OpenReport "PayProceeds_Cert", acViewReport, , , , SchNoX

   ' Exit Sub
    Dim myApp  As Word.Application
    Dim myDoc  As Word.Document
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputReport, "PayProceeds_Cert", acFormatPDF, "C:\FAAS\" & FaUserName & "\temp\PayProceeds_Cert.pdf"
    DoCmd.SetWarnings True
    Set myApp = CreateObject("Word.Application")
    Set myDoc = myApp.Documents.Add("C:\FAAS\" & FaUserName & "\temp\PayProceeds_Cert.pdf")
    myApp.Visible = True
End Sub

Private Sub PickModeToggle_Click()
    Dim i As Integer
    For i = Me.ScheduleList.ListCount - 1 To 0 Step -1
        Me.ScheduleList.Selected(i) = 0
    Next i
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From Pick_Schedule_WorkTable")
    DoCmd.SetWarnings True
    If Me.PickModeToggle = True Then
        Me.PickModeToggle.caption = ChrW(9660) & "Pick On"
        Me.GetSelectedUnits.Enabled = True
    Else
        Me.PickModeToggle.caption = "Pick Off"
        Me.GetSelectedUnits.Enabled = False
    End If
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch(s): "
End Sub

Public Sub ScheduleList_Click()
    Me.SchID = Me.ScheduleList
    Me.GroupID = Null
    Me.UnitID = Null
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch(s): "
    If Me.PickModeToggle = True Then
        DoCmd.SetWarnings False
        If Coalesce(Me.ScheduleList.Column(0), 0) > 0 Then DoCmd.RunSQL ("Insert into Pick_Schedule_WorkTable(SchID) SELECT " & Me.ScheduleList.Column(0) & ";")
        DoCmd.SetWarnings True
        Call ReselectPickedSchs
    Else
        Me.SchNo = Me.SchNo & " " & Me.ScheduleList.Column(1)
    End If
    SQLz = "SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType, MYear, UnitGroup " & _
        "FROM vw_SixKeys " & _
        "WHERE SchID=" & Me.ScheduleList & " ORDER BY AssetID;"
    Me.Units_List.RowSource = ""
    Me.PayOutList.RowSource = ""
    SQLz = ""
    If Me.PickModeToggle = False Then Call GetSelectedUnits_Click
    Call ClearSearch("All")
    Call ReselectPickedSchs
End Sub
Public Sub ReselectPickedSchs()
    Dim i, x As Integer
    Dim Spacer As String
    Me.SchNo = Me.ClientList.Column(1) & " MLNo: " & LeaseList.Column(1) & " Sch(s): "
    If Me.PickModeToggle = False Then Me.SchNo = Me.SchNo & " " & Me.ScheduleList.Column(1)
    SQLz = "Select SchID From Pick_Schedule_WorkTable Group By SchID Order by SchID Desc;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount > 0 Then
        For x = 1 To RST.RecordCount
            For i = Me.ScheduleList.ListCount - 1 To 0 Step -1
                If Me.ScheduleList.ItemData(i) = CStr(RST!SchID) Then
                    Me.ScheduleList.Selected(i) = 1
                    Spacer = ", "
                    If x = 1 Then Spacer = " "
                    If x > 1 And x = (RST.RecordCount) Then Spacer = " and "
                    Me.SchNo = Me.SchNo & Spacer & DLookup("SchNo", "Schedule", "SchID=" & RST!SchID)
                End If
            Next i
            RST.MoveNext
        Next x
    End If
    RST.close
    Set RST = Nothing
End Sub


Public Function ClearSearch(NotMe As String)
    If NotMe <> "SearchAsset" Then Me.SearchAsset = ""
    If NotMe <> "SearchUnit" Then Me.SearchUnit = ""
    If NotMe <> "SearchVin" Then Me.SearchVin = ""
    If NotMe <> "SearchStatus" Then Me.SearchStatus = ""
    If NotMe <> "SearchMake" Then Me.SearchMake = ""
    If NotMe <> "SearchType" Then Me.SearchType = ""
End Function
Private Sub SearchAsset_Change()
    Call ClearSearch("SearchAsset")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND AssetId Like '*" & Me.SearchAsset.Text & "*' ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz
    '  Call GetUnitCounts
End Sub

Private Sub SearchClient_Change()
    SQLz = "SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] " & _
        "From vw_SixKeys WHERE ClientShortName Like '*" & Me.SearchClient.Text & "*' Or CompanyName Like '*" & Me.SearchClient.Text & "*' " & _
        "GROUP BY clientID, ClientShortName, CompanyName " & _
        "ORDER BY CompanyName;"
    Me.ClientList.RowSource = SQLz
End Sub

Private Sub SearchUnit_Change()
    Call ClearSearch("SearchUnit")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND UnitNum Like '*" & Me.SearchUnit.Text & "*' ORDER BY UnitNum;"
    Me.Units_List.RowSource = SQLz
    'Call GetUnitCounts
End Sub
Private Sub SearchVIN_Change()
    Call ClearSearch("SearchVin")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND VIN Like '*" & Me.SearchVin.Text & "*' ORDER BY VIN;"
    Me.Units_List.RowSource = SQLz
    'Call GetUnitCounts
End Sub
Private Sub SearchStatus_Change()
    Call ClearSearch("SearchStatus")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND [Status] Like '*" & Me.SearchStatus.Text & "*' ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz
    ' Call GetUnitCounts
End Sub
Private Sub SearchMake_Change()
    Call ClearSearch("SearchMake")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND Make Like '*" & Me.SearchMake.Text & "*' ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz
    'Call GetUnitCounts
End Sub
Private Sub SearchType_Change()
    Call ClearSearch("SearchType")
    SQLz = Me.Units_List.RowSource
    SQLz = SplitString(SQLz, " ORDER BY", 0)
    SQLz = SplitString(SQLz, " AND", 0)
    SQLz = SQLz & " AND SubType Like '*" & Me.SearchType.Text & "*' ORDER BY AssetID;"
    Me.Units_List.RowSource = SQLz
    'Call GetUnitCounts
End Sub
Private Sub ClearLabel_Click()
    Me.SearchVin.SetFocus
    Call SearchVIN_Change
End Sub

Public Sub Units_List_Click()
    Me.UnitID = Me.Units_List.Column(2)
End Sub

Private Sub Units_List_DblClick(Cancel As Integer)
    On Error GoTo FixMe

FixMe:
End Sub
