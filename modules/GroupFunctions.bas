Attribute VB_Name = "GroupFunctions"
Option Compare Database
Public Function FinalEconomicsReport(SchIDx As Long, GroupIDx As Long, GrpNumx As String)
 Dim x, Ahx, SchUnits, Errorx As Integer
    Dim checkIT As String
    Dim Msgx As String
    Dim SQLz As String
    Msgx = ""

    Call MessageUser("Open", "Final Economics Report", "Getting data...")
    Ahx = Nz(DCount("sGrpID", "vw_SixKeys", "SchID=" & SchIDx & " AND UnitDescDocVer Like '*Auto Hauler*'"), 0)
    SchUnits = DCount("AssetID", "vw_SixKeys", "SchID=" & SchIDx & "")
   ' If Me.PVFeesYN = -1 Then
   '     checkIT = "Group Service Fee:       " & ValidateField("SGrpFSLPmt", "schgrp", , GroupIDX, SchIdx) & " (Calculate PV Service Fees is checked)" & vbNewLine
   '     If checkIT Like "*missing*" Then Msgx = Msgx & checkIT
   ' End If
    checkIT = "Pri Lease Term:" & ValidateField("LeaseTermPrimary", "schgrp", , GroupIDx, SchIDx) & vbNewLine
    If checkIT Like "*missing*" Then Msgx = Msgx & checkIT

    checkIT = "Equipment Cost:" & ValidateField("OrigEquipCost", "schgrp", , , SchIDx) & vbNewLine
    If checkIT Like "*missing*" Then Msgx = Msgx & checkIT

    checkIT = "Allowed Mileage:" & ValidateField("SGprAllow1", "schgrp", , GroupIDx, SchIDx) & vbNewLine
    If checkIT Like "*missing*" Then Msgx = Msgx & checkIT

    checkIT = "Unit Synd Rent:" & ValidateField("UNITRENT", "Units", , GroupIDx, SchIDx) & vbNewLine
    If checkIT Like "*missing*" And SchUnits <= Ahx Then Msgx = Msgx & checkIT

    checkIT = "Unit Synd RV:" & ValidateField("UnitSynResid", "Units", , GroupIDx, SchIDx) & vbNewLine
    If checkIT Like "*missing*" And SchUnits <= Ahx Then Msgx = Msgx & checkIT

    checkIT = "Unit Sch RV:" & ValidateField("UnitSalesResid", "Units", , GroupIDx, SchIDx) & vbNewLine
    If checkIT Like "*missing*" Then Msgx = Msgx & checkIT

    If Msgx Like "*missing*" Then
        Msgx = "Required data is missing that is preventing the report from being generated. Please check the fields below for missing data." & vbNewLine & vbNewLine & Msgx
        If DCount("sGrpID", "SchGrp", "sGrpSchID=" & SchIDx) = 1 Then
            Msgx = Msgx & vbNewLine & "Sorry, but the report 'may' not run. If you feel the need, contact IT."
            MsgBox Msgx, vbCritical, AtlaasVersion
            GoTo SingleReport
            Exit Function
        Else
            Msgx = Msgx & vbNewLine & "ATLAAS will try to open the 'single group' report for group: " & GrpNumx & "."
            MsgBox Msgx, vbExclamation, AtlaasVersion
            GoTo SingleReport
        End If
    End If
    On Error GoTo SingleReport
    Call MessageUser("Update", "Final Economics Report", "Running Report...")
    DoCmd.close acReport, "ScheduleEconomicsSyndReport"
    SQLz = "Select * FROM vw_EconoProfileGroup Where SchId=" & SchIDx
    DoCmd.OpenReport "ScheduleEconomicsSyndReport", acViewPreview, , , , SQLz
    Call MessageUser("Close")
    Exit Function
SingleReport:
    On Error GoTo -1
    On Error GoTo FixMe2
    If Not Msgx Like "*missing*" Then MsgBox "There was an Data error within the report " & vbNewLine & "There could be missing data within one of the other groups" & vbNewLine & "ATLAAS wil try to open the report for this group (" & GrpNumx & ")", vbInformation, AtlaasVersion
    DoCmd.close acReport, "ScheduleEconomicsSyndReport"
    SQLz = "Select * FROM vw_EconoProfileGroup Where GroupId=" & GroupIDx
    DoCmd.OpenReport "ScheduleEconomicsSyndReport", acViewPreview, , , , SQLz
    Call MessageUser("Close")
    Exit Function
FixMe2:
    Call MessageUser("Close")
    MsgBox "There was fatal data an error " & vbNewLine & "Most likely there is missing data within this group" & vbNewLine & "Sorry, the report 'may' not run, if you feel the need, contact IT", vbCritical, AtlaasVersion

End Function

Function SetUnitsInGroup(ClientIDx As Integer, FA_NotFA_All As String)
    Dim RST    As DAO.Recordset
    Dim x      As Integer
    Dim SQLz As String

    SQLz = SQLz & "SELECT clientID, ClientGroupId, MLOrig, SGrpID, Count(AssetID) AS Units "
    SQLz = SQLz & "From vw_SixKeys "
    If FA_NotFA_All = "FA" Then SQLz = SQLz & "WHERE SubType<>'Auto Hauler' "
    If FA_NotFA_All = "All" Then SQLz = SQLz & "WHERE SubType<>'Auto Hauler' "

    SQLz = SQLz & "GROUP BY clientID, ClientGroupId, MLOrig, SGrpID "
    SQLz = SQLz & "HAVING clientID=" & ClientIDx

    If FA_NotFA_All = "FA" Then SQLz = SQLz & " AND MLOrig = 'FA' "
    If FA_NotFA_All = "NotFA" Then SQLz = SQLz & " AND MLOrig <> 'FA' "

    SQLz = SQLz & ";"

    Set RST = CurrentDb.OpenRecordset(SQLz)
    If RST.RecordCount < 1 Then
    MsgBox "There are no update-able Groups for Client: " & DLookup("ClientShNm", "Clients", "ClientId=" & ClientIDx), vbInformation, AtlaasCoreVersion
    GoTo FixMe
    End If
    RST.MoveFirst
    MessageUser "Open", "Updating " & RST.RecordCount & " Groups", "Also testing OBC numbers"
    Do Until RST.EOF
        MessageUser "Update", "Updating " & RST.RecordCount - x & " Groups", "Also testing " & RST!UNITS & " OBC numbers"
        SQLz = "UPDATE SchGrp SET UnitsInGroup = " & RST!UNITS & " " & _
               "WHERE SGrpID =" & RST!SGrpID & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        RST.MoveNext
        x = x + 1
    Loop
    SQLz = "Update Units Set OBCUnitNumber = UnitUnitNum Where nz(OBCUnitNumber,'') = '' AND ClientId=" & ClientIDx
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
FixMe:
    MessageUser "Close"

End Function
