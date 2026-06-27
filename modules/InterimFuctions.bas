Attribute VB_Name = "InterimFuctions"
Option Compare Database

Dim SQLz As String
Dim xDraw, xBank, XLS As String
Public Function MakeDrawDownUnitReport(WhereField As String, WhereID As Long, xHeader As String)

    SQLz = "Select Client"
    SQLz = SQLz & ",Sch"
    SQLz = SQLz & ",Grp"
    SQLz = SQLz & ",AssetID"
    SQLz = SQLz & ",UnitNum"
    SQLz = SQLz & ",VIN As [VIN (partid)] "         ' Col-F
    SQLz = SQLz & ",ItemDesc"                       ' Col-G
    SQLz = SQLz & ",Vendor"
    SQLz = SQLz & ",InvoiceNum"
    SQLz = SQLz & ",VendorCost as [Vendor Cost]"    ' Col-J
    SQLz = SQLz & ",VendorAmt as [Pay Vendor]"      ' Col K
    SQLz = SQLz & ",FAAmount as [FA Amount]"        ' Col-L
    SQLz = SQLz & ",InterimAmt as [Interim Amount]"  ' Col-M
    SQLz = SQLz & ",DrawNumber as [Draw Num]"       ' Col-N
    SQLz = SQLz & ",DrawName"                  ' Col-O
    SQLz = SQLz & ",DateDrawn"                 ' Col-P
    SQLz = SQLz & ",VendorPaid"                ' Col-Q
    SQLz = SQLz & ",IFBank"                    ' Col-R
    SQLz = SQLz & ",FundedDate as [Synd Date]" ' Col-S
    SQLz = SQLz & ",PPCert"
    SQLz = SQLz & ",iType"
    SQLz = SQLz & " FROM vw_DrawDowns_Units "
    SQLz = SQLz & " WHERE " & WhereField & " = " & WhereID & " "
    SQLz = SQLz & " Order By DrawName,AssetID,IType DESC; "

    XLS = "|A,DD|C,TC|D,BC|D,TX|E,TC|F,TC|G,TC|J,$E|K,$E|L,$E|M,$E|N,BC|T,BC|V,BC|"
    XLS = XLS & "A,W15|B,W7|C,W7|J,W13|K,W13|L,W13|M,W13|N,W5|Q,W11|S,W13|A,Wrap|"
    Call MessageUser("Open", "Making Draw Down Unitds(Parts) Report", "This will be quick...")
    SendReport2Excel SQLz, "DrawDown", xHeader, True, XLS
    Call MessageUser("Close")
End Function

Public Function fixSetInterimBank()

SQLz = "Update Units Set UnitIFBankID=5649 " & _
" Where UnitIFBankID Is Null " & _
" AND UnitIFAmt > 0 " & _
" AND NOT DrawdownID IS Null"
  '  DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
SQLz = "Update DrawDowns " & _
"Set DrawBankId = x.UnitIFBankID " & _
"From (Select UnitIfBankID, D.DrawBankID,U.DrawdownId " & _
"     From Units As U " & _
"     Inner Join DrawDowns as D on U.DrawdownID=D.DrawdownID " & _
"     Where Not UnitIFBankID Is Null And Not u.DrawDownId Is Null " & _
"     AND UnitIfBankID <> D.DrawBankId " & _
"     Group By UnitIfBankID, U.DrawdownID, D.DrawBankId) as X " & _
"Where DrawDowns.DrawDownId = x.DrawDownId"
 '   DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Function
Public Function SetDrawIFDate(iIFDate As String)
    Dim RST As DAO.Recordset
    Dim xIfBankID As Long
    SQLz = "Select DrawDownID From Bulk_Update_WorkTable WHERE DrawDownID > 0 Group By DrawDownID"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    RST.MoveFirst
    If RST.RecordCount > 0 Then
        For i = 1 To RST.RecordCount
            Call SetInterimFundDates(RST!DrawDownId, iIFDate)
            xIfBankID = Nz(DMax("UnitIFBankID", "Units", "DrawDownID=" & RST!DrawDownId & ""), 0)
            If xIfBankID <> 0 Then Call SetInterimBank(RST!DrawDownId, xIfBankID)
            RST.MoveNext
        Next i
    End If
    RST.close
    Set RST = Nothing
End Function
Public Function SetDrawIFBank(iIFBankID As Long)
    Dim RST As DAO.Recordset
    Dim mIfDate As String
    SQLz = "Select DrawDownID From Bulk_Update_WorkTable WHERE DrawDownID > 0 Group By DrawDownID"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    RST.MoveFirst
    If RST.RecordCount > 0 Then
        For i = 1 To RST.RecordCount
            Call SetInterimBank(RST!DrawDownId, iIFBankID)
            mIfDate = Nz(DMax("UnitIFDate", "Units", "DrawDownID=" & RST!DrawDownId & ""), "1/1/1900")
            If mIfDate <> "1/1/1900" Then Call SetInterimFundDates(RST!DrawDownId, mIfDate)
            RST.MoveNext
        Next i
    End If
    RST.close
    Set RST = Nothing
End Function
Public Function SetInterimBank(iDrawDownID As Long, iInterimBankID As Long)
   Call MessageUser("Open", "One Sec, Getting Details", "Set Interim Bank...")

    If iInterimBankID = 0 Then iInterimBankID = 5649

    xDraw = DLookup("DrawName", "DrawDowns", "DrawDownId=" & iDrawDownID)
    xBank = DLookup("BankShortName", "Banks", "BankId=" & iInterimBankID)

    Call MessageUser("Update", "Set " & xBank & " for sibling Units", "Set Interim Bank for " & xDraw)
    SQLz = "UPDATE Units "
    SQLz = SQLz & "SET UnitIFBankID = " & iInterimBankID & " "
    SQLz = SQLz & "WHERE DrawDownID =" & iDrawDownID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)

    Call MessageUser("Update", "Seting " & xBank & " for sibling Parts", "Interim Bank for " & xDraw)
    SQLz = "UPDATE ChildParts "
    SQLz = SQLz & "SET cIFBankID = " & iInterimBankID & " "
    SQLz = SQLz & "WHERE DrawDownID =" & iDrawDownID & ";"
    DoCmd.RunSQL (SQLz)

    Call MessageUser("Update", "Setting " & xBank & " for the DRAW", "Updating " & xDraw)
    SQLz = "UPDATE DrawDowns "
    SQLz = SQLz & "SET DrawBankId = " & iInterimBankID & " "
    SQLz = SQLz & "WHERE DrawDownId=" & iDrawDownID & ";"
    DoCmd.RunSQL (SQLz)
    Call MessageUser("Update", "Updating the Forms", "Almost done... one second")
    On Error GoTo FixMe
    DoCmd.SetWarnings False
    Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    Forms!BulkChange_Form.BulkChange_ChildParts.Form.Requery
    DoCmd.SetWarnings True
    Call MessageUser("Close")
    Exit Function
FixMe:
    Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    Forms!BulkChange_Form.BulkChange_ChildParts.Form.Requery
    DoCmd.SetWarnings True
    Call MessageUser("Close")
End Function

Public Function SetInterimFundDates(iDrawDownID As Long, iInterimDate As String)
    xDraw = DLookup("DrawName", "DrawDowns", "DrawDownId=" & iDrawDownID)
    Call MessageUser("Open", "One Sec, Getting Details", "Set Interim DATE...")

    Call MessageUser("Update", "Set '" & iInterimDate & "' for sibling Units", "Set Interim DATE for " & xDraw)
    SQLz = "Update Units "
    If iInterimDate & "test" = "test" Then
        SQLz = SQLz & "Set UnitIFDate = NULL "
    Else
        SQLz = SQLz & "Set UnitIFDate = '" & iInterimDate & "' "
    End If
    SQLz = SQLz & "WHERE DrawDownID =" & iDrawDownID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)

    Call MessageUser("Update", "Set '" & iInterimDate & "' for sibling Parts", "Set Interim DATE for " & xDraw)
    SQLz = "Update ChildParts "
    If iInterimDate & "test" = "test" Then
        SQLz = SQLz & "Set cIFDate = NULL "
    Else
        SQLz = SQLz & "Set cIFDate = '" & iInterimDate & "' "
    End If
    SQLz = SQLz & "WHERE DrawDownID =" & iDrawDownID & ";"
    DoCmd.RunSQL (SQLz)

    Call MessageUser("Update", "Setting '" & iInterimDate & "' for the DRAW", "Updating " & xDraw)
    SQLz = "UPDATE DrawDowns "
    If iInterimDate & "test" = "test" Then
        SQLz = SQLz & "Set DateDrawn = NULL "
    Else
        SQLz = SQLz & "Set DateDrawn = '" & iInterimDate & "' "
    End If
    SQLz = SQLz & "WHERE DrawDownId=" & iDrawDownID & ";"
    DoCmd.RunSQL (SQLz)

    Call MessageUser("Update", "Updating the Forms", "Almost done... one second")
    On Error GoTo FixMe
    DoCmd.SetWarnings False
    Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    Forms!BulkChange_Form.BulkChange_ChildParts.Form.Requery
    DoCmd.SetWarnings True
    Call MessageUser("Close")
    Exit Function
FixMe:
    Forms!BulkChange_Form.BulkChange_Syndication.Form.Requery
    Forms!BulkChange_Form.BulkChange_ChildParts.Form.Requery
    DoCmd.SetWarnings True
    Call MessageUser("Close")
End Function
