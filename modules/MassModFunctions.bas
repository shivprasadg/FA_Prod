Attribute VB_Name = "MassModFunctions"
Option Compare Database
Option Explicit
Dim SQLz As String
Public Committ As Variant

Public Function CheckInterimFunded(PartDesc As String)
    Dim RST        As DAO.Recordset
    Dim subjectx As String
    Dim msg1 As String
    Dim EmailGroup As String
    SQLz = "SELECT Client,Schedule aS Sch, UnitGroup as Grp, Min(U.IFDAte) AS IFDate, Count(U.UnitID) as IFUnits, Sum(U.IFAmount) as IFAmount"
    SQLz = SQLz & " FROM ((vw_FiveKeys AS F"
    SQLz = SQLz & " INNER JOIN (SELECT UnitID, Sum(UnitIFAmt) As IFAmount, Min(UnitIfDate) AS IFDate, GroupRef"
    SQLz = SQLz & " FROM Units GROUP BY UnitID, GroupRef) As U ON F.GroupID = U.GroupREF)"
    SQLz = SQLz & " INNER JOIN  Bulk_Update_WorkTable AS W ON U.UnitID=W.UnitID) "
    SQLz = SQLz & " WHERE U.IFAmount > 0 "
    SQLz = SQLz & " Group BY Client,Schedule, UnitGroup;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount = 0 Then
        RST.close
        Set RST = Nothing
        Exit Function
    End If
    'make email message
    subjectx = "Child Part (" & PartDesc & ") Added after Draw! " & RST!Client & " | Sch " & RST!Sch & " | Grp " & RST!Grp
    msg1 = msg1 & "<html><body><H2 style=""color:RED;font-family:Segoe UI;""><b>Child Part Added on " & Date & " AFTER Setting Interim Funding</H2></b></p>"
    msg1 = msg1 & "<p style=""font-family:Calibri;font-size:11pt"">"
    msg1 = msg1 & "Re: " & subjectx & "<br/><br/>"
    msg1 = msg1 & "Client:  <B>" & RST!Client & "</B> Schedule: <B>" & RST!Sch & "</B> Group: <B>" & RST!Grp & "</B><br />"
    msg1 = msg1 & HSpaces(5) & "IF Date: <B>" & Nz(RST!IFDate, "tbd") & "</B><Br />"
    msg1 = msg1 & HSpaces(5) & "IF Amount: <B>" & Format(RST!IFAmount, "$#,#.#0") & "</B><Br />"
    msg1 = msg1 & HSpaces(5) & "IF Units: <B>" & RST!IFUnits & "</B><Br /><Br />"
    msg1 = msg1 & HSpaces(5) & "Part Added: <B> " & PartDesc & "</B><Br />"
    msg1 = msg1 & HSpaces(5) & "Date Added: <B> " & Date & "</B><Br />"
    msg1 = msg1 & HSpaces(5) & "Added By: <B>" & FaUserName & "</B><Br />"
    msg1 = msg1 & "<p style=""font-family:Segoe UI;font-size:8px""><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style=""font-family:Segoe UI;font-size:12px"">"
    msg1 = msg1 & "</body></html>"
    RST.close
    Set RST = Nothing
    EmailGroup = "SSweet@fleetadvantage.com"
    EmailGroup = EmailGroup & ";KDeStefano@fleetadvantage.com"
    EmailGroup = EmailGroup & ";DRamirez@fleetadvantage.com"
    EmailGroup = EmailGroup & ";SHora@fleetadvantage.com"
    EmailGroup = EmailGroup & ";RGulbranson@fleetadvantage.com"
    EmailGroup = EmailGroup & ";JCunningham@fleetadvantage.com"
    EmailGroup = EmailGroup & ";JRickette@fleetadvantage.com"
    EmailGroup = EmailGroup & ";MdeAguiar@fleetadvantage.com"
    EmailGroup = EmailGroup & ";KBabb@fleetadvantage.com"
    EmailGroup = EmailGroup & ";BAntonellis@FleetAdvantage.com"
    Call SendServerEmail(EmailGroup, "ljohnson@fleetadvantage.com", subjectx, msg1, "Operations@fleetadvantage.com")

End Function

Public Function MassConfirmIT(ByVal sFieldName As String, ByVal sNewValue As String) As String
    If sFieldName = "Rmkt.FABrkComm" Then sFieldName = "Bank's Commission"
    If sFieldName = "Rmkt.WhSValue" Then sFieldName = "FMVBase-Wholesale Value"
    If sFieldName = "Rmkt.BuyerRef" Then sFieldName = "BuyerID (the Unit's Buyer)"
    If sFieldName = "Rmkt.RmktSPRAdj" Then sFieldName = "Adjustment"
    If sFieldName = "Rmkt.RmktDeposit" Then sFieldName = "Deposit"
    If sFieldName = "Rmkt.RmktDepdt" Then sFieldName = "Deposit Paid Date"
    If sFieldName = "Rmkt.RmkSm" Then sFieldName = "Sales Person"
    If sFieldName = "Rmkt.Wholesale" Then sFieldName = "Sale Disposition"
    If sFieldName = "Rmkt.RmktBOSSent" Then sFieldName = "Date Bill Of Sale was Paid"
    If sFieldName = "Rmkt.RMKT.[SUF Sent]" Then sFieldName = "Suf Sent"
    If sFieldName = "Units.UnitYr" Then sFieldName = "Model Year"
    If sFieldName = "Units.UnitVendorPytDate" Then sFieldName = "Vendor Paid Date"
    If sFieldName = "Units.VendorRef" Then sFieldName = "VendorID (the Unit's Vendor)"
    If sFieldName Like "*UnitIF*" Then sFieldName = "Interim " & SplitString(sFieldName, "UnitIF", 1)
    If sNewValue & "test" = "test" Then
        MassConfirmIT = MsgBox("You're about set the " & sFieldName & " to Null, for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Update-Change confrim " & sFieldName & "?")
    Else
        MassConfirmIT = MsgBox("You're about set the " & sFieldName & " to " & sNewValue & vbNewLine & "for ALL units listed." & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, "Change " & sFieldName & "?")
    End If
End Function

Public Function UpdateRMKTtbl(uField As String, uValue As Variant, Optional uWhere As String, Optional uExtraField As String, Optional uFromName As String)
'===This is the RMKT Mass-Mod Update Routine =====
   If MassConfirmIT(uField, uValue) = vbNo Then Exit Function
    Dim uWorkTable As String
    uWorkTable = "RMKT_MMod_Work"
    If uFromName = "BuyersOrderMassMod" Then uWorkTable = "BuyersOrder_MMod_Work"
    SQLz = "UPDATE RMKT "
    SQLz = SQLz & "INNER JOIN " & uWorkTable & " "
    SQLz = SQLz & "ON RMKT.UnitRef = " & uWorkTable & ".UnitID " & vbCrLf
    Select Case VarType(uValue)
        Case vbInteger, vbLong, vbDouble, vbCurrency, vbBoolean:
                         SQLz = SQLz & "SET " & uField & " =" & uValue & " " & uExtraField            ' Integer, Long, Double, Currency, Boolean
        Case vbString:   SQLz = SQLz & "SET " & uField & " ='" & CStr(uValue) & "' " & uExtraField    ' String
        Case vbDate:     SQLz = SQLz & "SET " & uField & " =#" & CDate(uValue) & "# " & uExtraField   ' Date
        Case Else:       SQLz = SQLz & "SET " & uField & " =Null " & uExtraField                      ' NULL
    End Select
    SQLz = SQLz & " " & uWhere & ";"
    On Error Resume Next
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    On Error GoTo 0
    If Err.Number <> 0 Then MsgBox "UpdateRMKTtbl SQL Error: " & Err.Description
End Function

Public Function UpdateRUnitTbl(uField As String, uValue As Variant, Optional uWhere As String, Optional uExtraField As String, Optional uFromName As String)
'===This is the RMKT Mass-Mod Update For Units Table updates Routine =====
    If MassConfirmIT(uField, uValue) = vbNo Then Exit Function
    Dim uWorkTable As String
    uWorkTable = "RMKT_MMod_Work"
    If uFromName = "BuyersOrderMassMod" Then uWorkTable = "BuyersOrder_MMod_Work"
    SQLz = "UPDATE Units "
    SQLz = SQLz & "INNER JOIN " & uWorkTable & " "
    SQLz = SQLz & "ON Units.UnitID = " & uWorkTable & ".UnitID "
    Select Case VarType(uValue)
        Case vbInteger, vbLong, vbDouble, vbCurrency, vbBoolean:
                         SQLz = SQLz & "SET " & uField & " =" & uValue & " " & uExtraField            ' Integer, Long, Double, Currency, Boolean
        Case vbString:   SQLz = SQLz & "SET " & uField & " ='" & CStr(uValue) & "' " & uExtraField    ' String
        Case vbDate:     SQLz = SQLz & "SET " & uField & " =#" & CDate(uValue) & "# " & uExtraField   ' Date
        Case Else:       SQLz = SQLz & "SET " & uField & " =Null " & uExtraField                      ' NULL
    End Select
    SQLz = SQLz & " " & uWhere & ";"
    On Error GoTo FixMe
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    On Error GoTo 0
Exit Function
FixMe:
    On Error GoTo 0
    If Err.Number <> 0 Then MsgBox "UpdateRUnitTbl SQL Error: " & Err.Description
End Function

Public Function UpdateMassUnits(uField As String, uValue As Variant, Optional uWhere As String, Optional uExtraField As String)
'===This is the Units Mass-Mod Update Routine =====
    If MassConfirmIT(uField, uValue) = vbNo Then Exit Function
    SQLz = "UPDATE Units "
    SQLz = SQLz & "INNER JOIN Bulk_Update_WorkTable "
    SQLz = SQLz & "ON Units.UnitID = Bulk_Update_WorkTable.UnitID "
    Select Case VarType(uValue)
        Case vbInteger, vbLong, vbDouble, vbCurrency, vbBoolean:
                         SQLz = SQLz & "SET " & uField & " =" & uValue & " " & uExtraField           ' Integer, Long, Double, Currency, Boolean
        Case vbString:   SQLz = SQLz & "SET " & uField & " ='" & CStr(uValue) & "' " & uExtraField   ' String
        Case vbDate:     SQLz = SQLz & "SET " & uField & " =#" & CDate(uValue) & "# " & uExtraField  ' Date
        Case Else:       SQLz = SQLz & "SET " & uField & " =Null " & uExtraField                     ' NULL
    End Select
    SQLz = SQLz & " " & uWhere & ";"
    On Error Resume Next
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    On Error GoTo 0
    If Err.Number <> 0 Then MsgBox "UpdateMassUnits SQL Error: " & Err.Description
End Function

Public Function UpdateMassPARTS(uField As String, uValue As Variant, Optional uWhere As String, Optional uExtraField As String)
'===This is the ChildPArts Mass-Mod Update Routine =====
    If MassConfirmIT(uField, uValue) = vbNo Then Exit Function
    SQLz = "UPDATE ChildParts "
    SQLz = SQLz & "INNER JOIN Bulk_Update_WorkTable_Child "
    SQLz = SQLz & "ON ChildParts.ChildPartID = Bulk_Update_WorkTable_Child.ChildPartID "
    Select Case VarType(uValue)
        Case vbInteger, vbLong, vbDouble, vbCurrency, vbBoolean:
                         SQLz = SQLz & "SET " & uField & " =" & uValue & " " & uExtraField           ' Integer, Long, Double, Currency, Boolean
        Case vbString:   SQLz = SQLz & "SET " & uField & " ='" & CStr(uValue) & "' " & uExtraField   ' String
        Case vbDate:     SQLz = SQLz & "SET " & uField & " =#" & CDate(uValue) & "# " & uExtraField  ' Date
        Case Else:       SQLz = SQLz & "SET " & uField & " =Null " & uExtraField                     ' NULL
    End Select
    SQLz = SQLz & " " & uWhere & ";"
    On Error Resume Next
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    On Error GoTo 0
    If Err.Number <> 0 Then MsgBox "UpdateMassParts SQL Error: " & Err.Description
End Function
