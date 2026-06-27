Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String
Private Sub CloseMe_Click()
    DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
    Me.TitleCheckInSub.Visible = False
    Me.SetLeanRelease.Visible = False
    SQLz = "UPDATE DISTINCTROW FirstTitleLog " & _
           "INNER JOIN Units ON FirstTitleLog.UnitRef = Units.UnitID " & _
           "SET Units.UnitTitleRecd = [FirstTitleLog].[UnitTitleLogDate] WHERE (((Units.UnitTitleRecd) Is Null) AND ((FirstTitleLog.EntryTypeId)=1));"
    DoCmd.SetWarnings False
    'DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    SQLz = "SELECT ClientShortName AS Client, "
    SQLz = SQLz & " [Schedule] & ' | ' & [UnitGroup] AS SchGrp, UnitNum, VIN, LienRelease as Release, AssetID "
    SQLz = SQLz & " From vw_SixKeys "
    SQLz = SQLz & " INNER JOIN Units on vw_SixKeys.AssetId=Units.UnitID "
    SQLz = SQLz & " WHERE Not LienRelease is Null"
    SQLz = SQLz & " ORDER BY LienRelease Desc;"
    Me.LienReleaseList.RowSource = SQLz
End Sub

Private Sub LienReleasesRpt_Click()
    SQLz = "SELECT ClientShortName AS Client, "
    SQLz = SQLz & " [Schedule] & ' | ' & [UnitGroup] AS SchGrp, AssetID, UnitNum, VIN, UnitDescShortVer as UDesc, LienRelease as Release "
    SQLz = SQLz & " From vw_SixKeys "
    SQLz = SQLz & " INNER JOIN Units on vw_SixKeys.AssetId=Units.UnitID "
    SQLz = SQLz & " WHERE Not LienRelease is Null"
    SQLz = SQLz & " ORDER BY LienRelease Desc;"
    XLS = "|A,BL|B,TC|C,TX|D,TC|E,BC|G,BC|"
    SendReport2Excel SQLz, "Released", "Lien Release Entered", True, XLS, 0
End Sub

Private Sub LocVINBtn_Click()
    Dim Msgx, StatusX, Response, SQLi As String
    If IsNull(Me.seekvin) Then Msgx = Msgx & "    A VIN" & vbNewLine
    If IsNull(Me.seekMY) Then Msgx = Msgx & "    A Year" & vbNewLine
    If IsNull(Me.seekst) Then Msgx = Msgx & "    A State" & vbNewLine
    If Msgx <> "" Then
        MsgBox "You must input:" & vbNewLine & Msgx, , "Missing Data"
        Exit Sub
    End If
    StatusX = DLookup("UnitStatus", "Units", "UnitVIN='" & Me.seekvin & "'")
    If IsNull(StatusX) Then
        MsgBox "That VIN is not in ATLAAS Core" & vbNewLine & " Make sure it is the correct VIN", vbCritical, "VIN Not In System"
        Exit Sub
    End If
    Response = vbYes
    If StatusX = "T" Or StatusX = "O" Then 'Or StatusX = "R" Or StatusX = "PS" Then
        Response = MsgBox("You are attempting to add a Log for a Unit that may not belong to " & vbNewLine & "Fleet Advantage. This Unit may be a Purchase Lease Back" & vbNewLine & "Do you want to proceed?", vbQuestion + vbYesNo, "Not a FA unit")
    End If
    If Response = vbNo Then Exit Sub
    SQLi = "Select * From vw_SixKeys Where VIN='" & Me.seekvin & "' AND (Status <>'T' or Status <> 'O')"
    Dim RST As DAO.Recordset
    Set RST = CurrentDb.OpenRecordset(SQLi, dbOpenSnapshot)
    If RST.RecordCount < 1 Then Exit Sub
    Me.TitleCheckInSub.Visible = True
    Me.TitleCheckInSub.Form.UnitTitleLogDate = Date
    Me.TitleCheckInSub.Form.UnitID = RST!UnitID
    Me.TitleCheckInSub.Form.unitdesc = RST!UnitDescShortVer
    Me.TitleCheckInSub.Form.UnitVIN = RST!VIN
    Me.TitleCheckInSub.Form.UnitYr = RST!MYear
    Me.TitleCheckInSub.Form.SchGrp = RST!ClientShortName & " | " & RST!Schedule & " | " & RST!UnitGroup
    Me.TitleCheckInSub.Form.UnitTitleNumber = DLookup("UnitTitleNumber", "Units", "UnitVIN='" & Me.seekvin & "'")
    Me.SetLeanRelease.Visible = True
    If Not IsNull(DLookup("UnitTitleLocationRef", "Units", "UnitVIN='" & Me.seekvin & "'")) Then
        Me.TitleCheckInSub.Form.TitledState = DLookup("TitleSt", "UnitTitleHistoryLastEntry", "UnitRef=" & Me.TitleCheckInSub.Form.UnitID)
    Else
        Me.TitleCheckInSub.Form.TitledState = Me.seekst
    End If
    RST.close
    Set RST = Nothing

End Sub

Private Sub OpsReportExcel_Click()

    If Nz(Me.TitleLogList, 0) <> 0 Then
        Call MakeOPSReport(, Me.TitleLogList.Column(9))
    Else
        MsgBox "You MUST select a 'VIN'", vbCritical, AtlaasVersion
    End If
End Sub
Private Sub ScheduleCheck_Click()
    If Nz(Me.TitleSchList, 0) <> 0 Then
        Call MakeScheduleWorkTable(Me.TitleSchList)
        DoCmd.OpenReport "ScheduleCheckListNew", acViewPreview, , "SchID=" & Me.TitleSchList, acNormal
    Else
        MsgBox "You MUST select a 'Recent Log Entry Event'", vbCritical, AtlaasVersion
    End If
End Sub
Private Sub SeekVin_Change()

    SQLz = "SELECT UnitTitleHistoryLastEntry.UnitRef AS UnitID, UnitTitleHistoryLastEntry.UnitTitleLogDate AS LogDate, IIf([EntryTypeId]=1,'In','Out') AS [Action], " & _
           "Units.UnitVIN AS VIN, Units.UnitYr AS [Year], Nz([TitleSt],[TitledLocationState]) AS State, UnitTitleHistoryLastEntry.EntryReason AS Reason, " & _
           "UnitTitleHistoryLastEntry.TitleNumber, UnitTitleHistoryLastEntry.ExpectedReturn " & _
           "FROM UnitTitleHistoryLastEntry LEFT JOIN Units ON UnitTitleHistoryLastEntry.UnitRef = Units.UnitID " & _
           "WHERE Units.UnitVIN like '*" & Me.seekvin.Text & "*' ORDER BY UnitTitleHistoryLastEntry.UnitTitleLogDate DESC;"
    Me.TitleLogList.RowSource = SQLz
End Sub

Private Sub SetLeanRelease_Click()
    Dim Committ As String
    Committ = MsgBox("You are about to Lien Release Date for VIN:" & seekvin & vbNewLine & "Do you agree?", vbYesNo + vbQuestion, "Setting Lien Release")
    If Committ = vbNo Then Exit Sub
    SQLz = "Update Units Set LienRelease ='" & Date & "' Where UnitID=" & Me.TitleLogList.Column(0)
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Me.LienReleaseList.Requery
End Sub

Private Sub TitleLogList_Click()
    Me.Direction = Me.TitleLogList.Column(2)
    Me.seekvin = Me.TitleLogList.Column(3)
    Me.seekMY = Me.TitleLogList.Column(4)
    Me.seekst = Me.TitleLogList.Column(5)
    Me.SeekReason = Me.TitleLogList.Column(6)
    Me.TitleCheckInSub.Form.RecordSource = ""
    Me.SetLeanRelease.Visible = False
    Me.TitleCheckInSub.Visible = False


End Sub

Private Sub UpdateAllUnits_Click()
    SQLz = "UPDATE DISTINCTROW FirstTitleLog " & _
           "INNER JOIN Units ON FirstTitleLog.UnitRef = Units.UnitID " & _
           "SET Units.UnitTitleRecd = [FirstTitleLog].[UnitTitleLogDate] WHERE (((Units.UnitTitleRecd) Is Null) AND ((FirstTitleLog.EntryTypeId)=1));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Sub
