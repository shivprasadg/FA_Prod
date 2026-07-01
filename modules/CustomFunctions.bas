Attribute VB_Name = "CustomFunctions"
Option Compare Database
Option Explicit
Dim SQLz As String
'Public Function WebView_Exec(web As WebView2, js As String) As String
'    Dim task As Object
'    Set task = web.ExecuteScriptAsync(js)
 '   Do While task.Status = 0 ' MKT: 0 = running
 '       DoEvents
 '   Loop

  '  WebView_Exec = task.result
'End Function

Public Function SplitString(SplitToString As String, DelimiterStr As String, Section As Integer)
    On Error GoTo Handler
    SplitString = Split(SplitToString, DelimiterStr)(Section)
    Exit Function
Handler:
    If InStr(1, SplitToString, DelimiterStr) < Section Then
        SplitString = "Error: (" & Err.Number & ") " & Err.Description & " Section Excceded"
        Exit Function
    End If
    SplitString = "Error: (" & Err.Number & ") " & Err.Description
End Function
Public Function IsAutohauler(LookUpField As String, LookupID As Long) As Integer
    IsAutohauler = DCount("AssetID", "vw_SixKeys", "SubType ='Auto Hauler' AND " & LookUpField & "=" & LookupID & "")
End Function
Function DoesMasterLeaseExist(MLNox As String, ClientIDx As Integer) As Boolean
    Dim RST As DAO.Recordset
    DoesMasterLeaseExist = False
    SQLz = "SELECT MLNo,Client FROM MstrLease WHERE Client=" & ClientIDx & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then Exit Function
    RST.MoveFirst
    Do Until RST.EOF
        If MLNox = RST!MLNo Then
            DoesMasterLeaseExist = True
            Exit Do
        End If
        RST.MoveNext
    Loop
    RST.close
    Set RST = Nothing
End Function
Function DoesScheduleExist(SchNoX As String, MLIDx As Integer) As Boolean
    Dim RST As DAO.Recordset
    DoesScheduleExist = False
    SQLz = "SELECT SchNo,SchMLRef FROM Schedule WHERE SchMLRef=" & MLIDx & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then Exit Function
    RST.MoveFirst
    Do Until RST.EOF
        If SchNoX = RST!SchNo Then
            DoesScheduleExist = True
            Exit Do
        End If
        RST.MoveNext
    Loop
    RST.close
    Set RST = Nothing
End Function
Function ListSchedules(MLIDx As Integer) As String
    Dim RST As DAO.Recordset
    ListSchedules = ""
    SQLz = "SELECT SchNo,SchMLRef, DateSchAdded FROM Schedule WHERE SchMLRef=" & MLIDx & " Order By SchNo Desc;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then Exit Function
    RST.MoveFirst
    Do Until RST.EOF
        ListSchedules = ListSchedules & "     Schedule " & Chr(187) & " " & RST!SchNo & "  -  added in @ " & RST!DateSchAdded & vbNewLine
        RST.MoveNext
    Loop
    RST.close
    Set RST = Nothing
End Function
Function DoesGroupExist(GroupName As String, SchIDx As Integer) As Boolean
    Dim RST As DAO.Recordset
    DoesGroupExist = False
    SQLz = "SELECT SGrpNo,SGrpSchID FROM SchGrp WHERE SGrpSchID=" & SchIDx & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then Exit Function
    RST.MoveFirst
    Do Until RST.EOF
        If GroupName = RST!SGrpNO Then
            DoesGroupExist = True
            Exit Do
        End If
        RST.MoveNext
    Loop
    RST.close
    Set RST = Nothing
End Function
Function IsAlphaNumeric(pValue) As Boolean
    Dim LPos As Integer
    Dim LChar As String
    Dim LValid_Values As String
    'Start at first character in pValue
    LPos = 1
    'Set up values that are considered to be alphanumeric
    LValid_Values = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ+-.0123456789"
    'Test each character in pValue
    While LPos <= Len(pValue)

        'Single character in pValue
        LChar = Mid(pValue, LPos, 1)
        'If character is not alphanumeric, return FALSE
        If InStr(LValid_Values, LChar) = 0 Then
            IsAlphaNumeric = False
            Exit Function
        End If
        'Increment counter
        LPos = LPos + 1
    Wend
    'Value is alphanumeric, return TRUE
    IsAlphaNumeric = True
End Function
Public Function FormDocumentor(FormX As String, Optional Closeit As Boolean)
    Dim ctl As Control
    DoCmd.OpenForm (FormX)
    Debug.Print "Form: " & FormX
    Debug.Print "Label|Source"
    On Error Resume Next
    For Each ctl In Forms(FormX)
        Debug.Print ctl.Controls.Item(0).caption & "|" & ctl.ControlSource
    Next ctl
    If Closeit = True Then DoCmd.close acForm, FormX
End Function
Public Function ConcatRelated(strField As String, _
                              strTable As String, _
                              Optional strWhere As String, _
                              Optional strOrderBy As String, _
                              Optional strSeparator = ", ", _
                              Optional str2ndSeparator = "[", _
                              Optional str3rdSeparator = "]") As Variant
    On Error GoTo Err_Handler
    'Purpose:   Generate a concatenated string of related records.
    'Return:    String variant, or Null if no matches.
    'Arguments: strField = name of field to get results from and concatenate.
    '           strTable = name of a table or query.
    '           strWhere = WHERE clause to choose the right values.
    '           strOrderBy = ORDER BY clause, for sorting the values.
    '           strSeparator = characters to use between the concatenated values.
    'Notes:     1. Use square brackets around field/table names with spaces or odd characters.
    '           2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
    '           3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
    '           4. Returning more than 255 characters to a recordset triggers this Access bug:
    '               http://allenbrowne.com/bug-16.html
    Dim rs As DAO.Recordset     'Related records
    Dim rsMV As DAO.Recordset     'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.

    'Initialize to Null
    ConcatRelated = Null

    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).Type > 100)

    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).Value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
            If HowManyInStr(strField, ",") > 0 Then strOut = strOut & str2ndSeparator & rs(1) & " "
            If HowManyInStr(strField, ",") > 1 Then strOut = strOut & rs(2)
            If HowManyInStr(strField, ",") > 2 Then strOut = strOut & rs(3)
            If HowManyInStr(strField, ",") > 0 Then strOut = strOut & str3rdSeparator
            strOut = strOut & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.close

    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Handler:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ConcatRelated()"
    Resume Exit_Handler
End Function

Public Function ParaPass(Queryx As String, ParaField1 As String, ParaValue1 As String _
     , Optional ParaField2 As String, Optional ParaValue2 As String _
     , Optional ParaField3 As String, Optional ParaValue3 As String _
     , Optional ParaField4 As String, Optional ParaValue4 As String _
     , Optional ParaField5 As String, Optional ParaValue5 As String _
                                    )
    Dim qdf As DAO.QueryDef
    Set qdf = CurrentDb.QueryDefs(Queryx)
    qdf.Parameters(ParaField1).Value = ParaValue1
    qdf!ParaField1 = ParaValue1
    DoCmd.SetWarnings False
    qdf.Execute
    DoCmd.SetWarnings True
    Set qdf = Nothing
End Function

'
'Private Enum CaseTypes
'    CaseTypeSimple = 1
'    CaseTypeSearched = 2
'End Enum
Public Function PrintReport(strReport As String) As Boolean
' open report in PREVIEW mode but HIDDEN
    DoCmd.OpenReport strReport, acViewPreview, , , acHidden
    ' open the dialog form to let the user choose printing options
    DoCmd.OpenForm "dlgPrinter", , , , , acDialog, strReport
    With Forms!dlgPrinter
        If .Tag <> "Cancel" Then
            Set Reports(strReport).Printer = Application.Printers((!cmbPrinter))
            Reports(strReport).Printer.Orientation = !optLayout
            Application.Echo False
            DoCmd.SelectObject acReport, strReport
            DoCmd.PrintOut acPages, !txtPageFrom, !txtPageTo
            PrintReport = True
        End If
    End With
    DoCmd.close acForm, "dlgPrinter"
    DoCmd.close acReport, strReport
    Application.Echo True
End Function
Function ProperCase(ByVal strOneLine As String, intChangeType As Integer) As String
'---------------------------------------------------------------
'- This function will convert a string to Proper Case -
'- The initial letter of each word is capitalised. -
'- It will also handle special names such as O', Mc and -
'- hyphenated names -
'- if intChangeType = 1, all text is converted to proper case. -
'- e.g. 'FRED' is converted to 'Fred' -
'- if intChangeType = 0, upper case text is not converted. -
'- e.g. 'fred' becomes 'Fred', but 'FRED' remains unchanged. -
'---------------------------------------------------------------
    Dim i As Integer
    Dim bChangeFlag As Boolean
    Dim strResult As String
    '----------------------------------------------------------
    '- No characters in string - nothing to do -
    '----------------------------------------------------------
    If Len(strOneLine) = 0 Then
        ProperCase = ""
        Exit Function
    End If
    '----------------------------------------------------------
    '- Always set first letter to upper case -
    '----------------------------------------------------------
    strResult = UCase$(Left$(strOneLine, 1))
    '----------------------------------------------------------
    '- Now look at the rest of the string -
    '----------------------------------------------------------
    For i = 2 To Len(strOneLine)

        '----------------------------------------------------------
        '- If the previous letter triggered a capital, change -
        '- this letter to upper case -
        '----------------------------------------------------------
        If bChangeFlag = True Then
            strResult = strResult & UCase$(Mid$(strOneLine, i, 1))
            bChangeFlag = False
            '----------------------------------------------------------
            '- In other cases change letter to lower case if required -
            '----------------------------------------------------------
        Else
            If intChangeType = 1 Then
                strResult = strResult & LCase$(Mid$(strOneLine, i, 1))
            Else
                strResult = strResult & Mid$(strOneLine, i, 1)
            End If
        End If

        '----------------------------------------------------------
        '- Set change flag if a space, apostrophe or hyphen found -
        '----------------------------------------------------------
        Select Case Mid$(strOneLine, i, 1)
        Case " ", "'", "-"
            bChangeFlag = True
        Case Else
            bChangeFlag = False
        End Select
    Next i
    '----------------------------------------------------------
    '- Special handling for Mc at start of a name -
    '----------------------------------------------------------
    If Left$(strResult, 2) = "Mc" Then
        Mid$(strResult, 3, 1) = UCase$(Mid$(strResult, 3, 1))
    End If

    i = InStr(strResult, " Mc")
    If i > 0 Then
        Mid$(strResult, i + 3, 1) = UCase$(Mid$(strResult, i + 3, 1))
    End If

    '----------------------------------------------------------
    '- Special handling for Mac at start of a name -
    '----------------------------------------------------------
    If Left$(strResult, 3) = "Mac" And Len(strResult) > 3 Then
        Mid$(strResult, 4, 1) = UCase$(Mid$(strResult, 4, 1))
    End If

    i = InStr(strResult, " Mac")
    If i > 0 Then
        Mid$(strResult, i + 4, 1) = UCase$(Mid$(strResult, i + 4, 1))
    End If

    ProperCase = strResult
End Function
Function SplitRef(SplitString As String, Section As Integer)
Dim StartPos As Integer
Dim EndPos As Integer
    Select Case Section
    Case 1:
        StartPos = 1
        EndPos = InStr(SplitString, "-") - 2
    Case 2:
        StartPos = InStr(SplitString, "-") + 2
        EndPos = InStr(SplitString, ",") - 1
    Case 3:
        StartPos = InStr(SplitString, ",") + 2
        EndPos = InStrRev(SplitString, " ") - 1
    Case 4:
        StartPos = InStrRev(SplitString, " ") + 1
        EndPos = Len(SplitString)
    Case 5:
        StartPos = 1
        EndPos = InStr(SplitString, " (") - 1
    End Select
    SplitRef = Mid(SplitString, StartPos, EndPos - StartPos + 1)
End Function

Public Function WriteStringToFile(TextFileName As String, StringToWrite As String)
    Dim intFileNum As Integer
    If FileExists(TextFileName) Then Kill TextFileName
    intFileNum = FreeFile
    ' change Output to Append if you want to add to an existing file
    ' rather than creating a new file each time
    Open TextFileName For Output As intFileNum
    Print #intFileNum, StringToWrite
    Close intFileNum
End Function

Public Function MakeScheduleReport(SchIDx As Integer, Optional TwoTab As Boolean)
    Call MessageUser("Open", "Creating 'Docs' Report", "..with ATF Detail on 2nd tab")
    Dim SQLt, Cost1X, Cost2X, XLS, TitleX As String
    Dim intRow, intCol, x, Countx As Integer
    Dim RST As DAO.Recordset
    Dim qdf As DAO.QueryDef
    On Error GoTo FixMe
    On Error Resume Next
    Set RST = CurrentDb.OpenRecordset("SELECT * FROM Cost_ChildParts_Types0 WHERE SchID=" & SchIDx & " ORDER BY Cpair")
    RST.MoveFirst
    RST.MoveLast
    RST.MoveFirst
    intRow = RST.RecordCount
    intCol = (intRow * 5) + 1
    ReDim ChildType(1 To intRow) As String
    For x = 1 To intRow
        ChildType(x) = RST!Cpair
        RST.MoveNext
    Next x
    RST.close
    Set RST = Nothing
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "Cost_ChildParts_AssetID"
    DoCmd.DeleteObject acQuery, "Cost_ChildParts_Pivot_Vendors"
    DoCmd.DeleteObject acQuery, "Cost_ChildParts_Pivot_Cost"
    DoCmd.DeleteObject acQuery, "Cost_ChildParts_Pivot_InvoiceNum"
    DoCmd.DeleteObject acQuery, "Cost_ChildParts_Pivot_InvoiceDate"
    DoCmd.DeleteObject acQuery, "Cost_ChildParts_Pivot_Make"
    DoCmd.DeleteObject acQuery, "Cost_ChildParts_Pivot_DrawNum"
    DoCmd.DeleteObject acQuery, "Cost_ChildParts_AllTypes"
    DoCmd.DeleteObject acQuery, "SchDocsReportFinal"
    DoCmd.DeleteObject acQuery, "SchDocsVendorUnion"
    'Make ChildParts_AssetID
    SQLt = "Select AssetID From Cost_ChildParts " & _
           "WHERE (((Cost_ChildParts.SchID) = " & SchIDx & ")) " & _
           "GROUP BY Cost_ChildParts.AssetId;"
    Set qdf = CurrentDb.CreateQueryDef("Cost_ChildParts_AssetID", SQLt)
    'If intRow < 1 Then GoTo MakeReport
    On Error GoTo FixMe

    'Make Vendor Pivot
    SQLt = "TRANSFORM First(Vendors.VendorSNm) AS ChildVendor " & "SELECT Cost_ChildParts.AssetID " & _
           "FROM Cost_ChildParts INNER JOIN Vendors ON Cost_ChildParts.ChildVendor = Vendors.VendorID " & _
           "WHERE (((Cost_ChildParts.SchID)=" & SchIDx & ")) " & "GROUP BY Cost_ChildParts.AssetID PIVOT [PartType] & ' Vendor';"
    Set qdf = CurrentDb.CreateQueryDef("Cost_ChildParts_Pivot_Vendors", SQLt)
    Call MessageUser("Open")
    Call MessageUser("Update", "Making supporting pivots", "...just a bit more...")

    'Make Cost Pivot
    SQLt = "TRANSFORM First(Cost_ChildParts.PartCost) AS FirstOfPartCost " & "SELECT Cost_ChildParts.AssetID " & _
           "FROM Cost_ChildParts " & "WHERE (((Cost_ChildParts.SchID)=" & SchIDx & ")) " & _
           "GROUP BY Cost_ChildParts.AssetID " & "PIVOT [PartClass] & Chr (13) & Chr (10) & [PartType] & ' Cost';"
    Set qdf = CurrentDb.CreateQueryDef("Cost_ChildParts_Pivot_Cost", SQLt)

    'Make SchDocVendorUnion
    SQLt = "SELECT UnitID, UnitVendor as VendorName, UnitCost as Cost, UnitInvoiceNum as InvoiceNum, UnitType, VendorID " & _
           "FROM SchDocsVendorUnit WHERE SchID=" & SchIDx & " " & _
           "UNION ALL Select UnitID, ChildVendor, ChildCost,ChildInvoiceNum, ChildType, ChildVendorID " & _
           "FROM SchDocsVendorChild WHERE SchID=" & SchIDx & " " & _
           "ORDER BY UnitId, Cost DESC , UnitType;"
    Set qdf = CurrentDb.CreateQueryDef("SchDocsVendorUnion", SQLt)

    'Make InvoiceNum Pivot
    SQLt = "TRANSFORM First(Cost_ChildParts.PartInvoice) AS FirstOfPartInvoice " & _
           "SELECT Cost_ChildParts.AssetID " & "FROM Cost_ChildParts " & _
           "WHERE (((Cost_ChildParts.SchID)=" & SchIDx & ")) " & _
           "GROUP BY Cost_ChildParts.AssetID " & _
           "PIVOT [PartClass] & Chr (13) & Chr (10) & [PartType] & ' Invoice';"
    Set qdf = CurrentDb.CreateQueryDef("Cost_ChildParts_Pivot_InvoiceNum", SQLt)

    'Make InvoiceDate Pivot
    SQLt = "TRANSFORM First(Cost_ChildParts.PartPaidDate) AS FirstOfPartPaidDate " & _
           "SELECT Cost_ChildParts.AssetID " & "FROM Cost_ChildParts " & _
           "WHERE (((Cost_ChildParts.SchID)=" & SchIDx & ")) " & _
           "GROUP BY Cost_ChildParts.AssetID " & _
           "PIVOT [PartClass] & Chr (13) & Chr (10) & [PartType] & ' Paid';"
    Set qdf = CurrentDb.CreateQueryDef("Cost_ChildParts_Pivot_InvoiceDate", SQLt)

   'Make PartMake Pivot
    SQLt = "TRANSFORM First(Cost_ChildParts.PartMake) AS FirstOfPartMake " & _
           "SELECT Cost_ChildParts.AssetID " & "FROM Cost_ChildParts " & _
           "WHERE (((Cost_ChildParts.SchID)=" & SchIDx & ")) " & _
           "GROUP BY Cost_ChildParts.AssetID " & _
           "PIVOT [PartClass] & Chr (13) & Chr (10) & [PartType] & ' Make';"
    Set qdf = CurrentDb.CreateQueryDef("Cost_ChildParts_Pivot_Make", SQLt)

    'Make PartDraw Pivot
    SQLt = "TRANSFORM First(Cost_ChildParts.PartDraw) AS PartDraw " & _
           "SELECT Cost_ChildParts.AssetID " & "FROM Cost_ChildParts " & _
           "WHERE (((Cost_ChildParts.SchID)=" & SchIDx & ")) " & _
           "GROUP BY Cost_ChildParts.AssetID " & _
           "PIVOT [PartClass] & Chr (13) & Chr (10) & [PartType] & ' DrawNum';"
    Set qdf = CurrentDb.CreateQueryDef("Cost_ChildParts_Pivot_DrawNum", SQLt)


    If intRow > 0 Then
        SQLt = "SELECT Cost_ChildParts_AssetID.AssetId, "
        For x = 1 To intRow
            Cost1X = Cost1X & " Cost_ChildParts_Pivot_Make.[" & ChildType(x) & " Make], "
            Cost1X = Cost1X & " Cost_ChildParts_Pivot_Cost.[" & ChildType(x) & " Cost], "
            Cost1X = Cost1X & " Cost_ChildParts_Pivot_InvoiceNum.[" & ChildType(x) & " Invoice], "
            Cost1X = Cost1X & " Cost_ChildParts_Pivot_InvoiceDate.[" & ChildType(x) & " Paid], "
            Cost1X = Cost1X & " Cost_ChildParts_Pivot_DrawNum.[" & ChildType(x) & " DrawNum], "
            Cost2X = Cost2X & "Cdbl(nz([" & ChildType(x) & " Cost],0))+"
        Next x
        ' If Right(SQLT, 2) = ", " Then SQLT = Left(SQLT, Len(SQLT) - 2) & " "
        Cost2X = Left(Cost2X, Len(Cost2X) - 1) & " As [ChildPartTotal] "
        SQLt = SQLt & Cost1X & Cost2X
    Else
        SQLt = "SELECT Cost_ChildParts_AssetID.AssetId "
    End If
MakeReport:
    On Error GoTo FixMe

    SQLt = SQLt & "FROM ((((Cost_ChildParts_AssetID " & _
    "LEFT JOIN Cost_ChildParts_Pivot_Make ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_Make.AssetID) " & _
    "LEFT JOIN Cost_ChildParts_Pivot_Cost ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_Cost.AssetID) " & _
    "LEFT JOIN Cost_ChildParts_Pivot_InvoiceNum ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_InvoiceNum.AssetID) " & _
    "LEFT JOIN Cost_ChildParts_Pivot_InvoiceDate ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_InvoiceDate.AssetID)  " & _
    "LEFT JOIN Cost_ChildParts_Pivot_DrawNum ON Cost_ChildParts_AssetID.AssetId = Cost_ChildParts_Pivot_DrawNum.AssetID;"

    Set qdf = CurrentDb.CreateQueryDef("Cost_ChildParts_AllTypes", SQLt)

    SQLt = "SELECT SU.[Lease No]"
    SQLt = SQLt & ",SU.Lessee"
    SQLt = SQLt & ",SU.Co_Lessee"
    SQLt = SQLt & ",SU.SchNo"
    SQLt = SQLt & ",SU.Group"
    SQLt = SQLt & ",SU.[Unit No]"
    SQLt = SQLt & ",SU.Unitvin"
    SQLt = SQLt & ",SU.Desc"
    SQLt = SQLt & ",SU.UnitCost"
    SQLt = SQLt & ",SU.VendorCost "
    SQLt = SQLt & ",SU.UnitInvoiceNum As [Invoice Num]"
    SQLt = SQLt & ",SU.UnitInvoiceDate As [Invoice Date]"
    SQLt = SQLt & ",SU.UnitInvoicePaid As [Invoice Paid]"
    SQLt = SQLt & ", "
    If intRow > 0 Then SQLt = SQLt & Cost1X & Cost2X & ", "
    If intRow > 0 Then SQLt = SQLt & "[UnitCost]+[ChildPartTotal] AS [Unit OEC]"
    If intRow < 1 Then SQLt = SQLt & "0 as [ChildPartTotal], "
    If intRow < 1 Then SQLt = SQLt & "[UnitCost]+0 AS [Unit OEC]"
    'SQLt = SQLt & " [UnitCost]+[ChildPartTotal] AS UnitTotalCost"
    SQLt = SQLt & ",SU.AssetId"
    SQLt = SQLt & ",SU.Status"
    SQLt = SQLt & ",SU.FAID"
    SQLt = SQLt & ",SU.EstiDeliveryDate as [Est Delv]"
    SQLt = SQLt & ",SU.ActualDelDate as [Actual Delv]"
    SQLt = SQLt & ",SU.DeliveryLoc"
    SQLt = SQLt & ",SU.InService"
    SQLt = SQLt & ",SU.[Title Num]"
    SQLt = SQLt & ",SU.TitledState"
    SQLt = SQLt & ",SU.Location As Domicile"
    SQLt = SQLt & ",SU.StreetAddress as Address"
    SQLt = SQLt & ",SU.City"
    SQLt = SQLt & ",SU.State"
    SQLt = SQLt & ",SU.Zip"
    SQLt = SQLt & ",SU.County"
    SQLt = SQLt & ",SU.InterimBank"
    SQLt = SQLt & ",SU.InterimDate AS [Interim Funded]"
    SQLt = SQLt & ",SU.InterimAmt"
    SQLt = SQLt & ",SU.DrawName"
    SQLt = SQLt & ",SU.CustodianID "
    SQLt = SQLt & ",SU.TaxExempt as [Tax Exempt]"
    SQLt = SQLt & ",SU.UpFront As [Up Front]"
    SQLt = SQLt & ",SU.UnitTaxNotes "
    SQLt = SQLt & ",SU.[State Tax]"
    SQLt = SQLt & ",SU.[County Tax]"
    SQLt = SQLt & ",SU.[City Tax]"
    SQLt = SQLt & ",SU.[Other Tax]"
    SQLt = SQLt & ",SU.[Loc Tax] AS [Total Tax]"
    SQLt = SQLt & ",SU.LTaxRate AS [ATLAAS TaxRate]"
    SQLt = SQLt & "FROM ScheduleUnitDocs3 As SU "
    SQLt = SQLt & "LEFT JOIN Cost_ChildParts_AllTypes ON SU.AssetId = Cost_ChildParts_AllTypes.AssetId "
    SQLt = SQLt & "WHERE (((SU.AssetID) Is Not Null)) "
    SQLt = SQLt & "ORDER BY SU.Group, SU.[Unit No], SU.AssetId;"

'  If intRow < 1 Then SQLt = Replace(SQLt, "[UnitCost]+[ChildPartTotal]", "[UnitCost]+0")

    Set qdf = CurrentDb.CreateQueryDef("SchDocsReportFinal", SQLt)

    XLS = "|A,DD4|D,BC|E,BC|F,TC|I,$E|J,$E|K,TC|G,CX|L,TC|BN,TC|BN,W06|BO,TC|BO,W06|"    'AB,$E|"
    x = 14
    If intRow > 0 Then
        Dim R1 As Integer
        R1 = 170 'Grey color start
        '  XLS = XLS & "N,TC|"
        For x = 14 To (intCol + 8) Step 5
            R1 = R1 - x
            If R1 < 0 Then R1 = 150
            XLS = XLS & ColumnLetter((x)) & ",TC|"
            XLS = XLS & ColumnLetter((x + 1)) & ",$E|" & ColumnLetter((x)) & ",cColor:3:" & R1 & "." & R1 & "." & R1 & "|"
            XLS = XLS & ColumnLetter((x + 1)) & ",cColor:3:" & R1 & "." & R1 & "." & R1 & "|"
            XLS = XLS & ColumnLetter((x + 1)) & ",W14|"
            XLS = XLS & ColumnLetter((x + 2)) & ",cColor:3:" & R1 & "." & R1 & "." & R1 & "|"
            XLS = XLS & ColumnLetter((x + 2)) & ",TC|"
            XLS = XLS & ColumnLetter((x + 3)) & ",cColor:3:" & R1 & "." & R1 & "." & R1 & "|"
            XLS = XLS & ColumnLetter((x + 3)) & ",TC|"
            XLS = XLS & ColumnLetter((x + 4)) & ",cColor:3:" & R1 & "." & R1 & "." & R1 & "|"
            XLS = XLS & ColumnLetter((x + 4)) & ",TC|"
            XLS = XLS & ColumnLetter((x + 4)) & ",W11|"
        Next x
    End If
    XLS = XLS & ColumnLetter((x)) & ",$E|" & ColumnLetter((x)) & ",W12|" & ColumnLetter((x)) & ",cColor:3:0.0.40|"  ' [ChildPartTotal]
    XLS = XLS & ColumnLetter((x + 1)) & ",$E|" & ColumnLetter((x + 1)) & ",cColor:3:60.125.34|"   ' [Unit OEC]
    XLS = XLS & ColumnLetter((x + 2)) & ",BC|"   'AssetID
    XLS = XLS & ColumnLetter((x + 3)) & ",TC|"   'UnitStatus
    XLS = XLS & ColumnLetter((x + 4)) & ",TC|" & ColumnLetter((x + 4)) & ",cColor:3:0.0.0|" 'FAID
    XLS = XLS & ColumnLetter((x + 10)) & ",TC|"  'TitleState
    XLS = XLS & ColumnLetter((x + 14)) & ",TC|"  'LocState
    XLS = XLS & ColumnLetter((x + 17)) & ",TC|" & ColumnLetter((x + 17)) & ",cColor:3:12.118.158|"  'InterimBank
    XLS = XLS & ColumnLetter((x + 18)) & ",W11|" & ColumnLetter((x + 18)) & ",cColor:3:12.118.158|" '[Interim Funded]
    XLS = XLS & ColumnLetter((x + 19)) & ",$E|" & ColumnLetter((x + 19)) & ",cColor:3:12.118.158|" & ColumnLetter((x + 18)) & ",W12|" 'InterimAmt
    XLS = XLS & ColumnLetter((x + 20)) & ",TC|" & ColumnLetter((x + 20)) & ",cColor:3:12.118.158|" 'DrawName
    XLS = XLS & ColumnLetter((x + 21)) & ",TC|"   'CustodianID

    XLS = XLS & ColumnLetter((x + 25)) & ",TC|" & ColumnLetter((x + 25)) & ",W8|" ' State Tax
    XLS = XLS & ColumnLetter((x + 26)) & ",TC|" & ColumnLetter((x + 26)) & ",W8|" ' County Tax"
    XLS = XLS & ColumnLetter((x + 27)) & ",TC|" & ColumnLetter((x + 27)) & ",W8|" ' City Tax"

    XLS = XLS & ColumnLetter((x + 29)) & ",P3|" & ColumnLetter((x + 28)) & ",cColor:3:red|" 'OtherTax
    XLS = XLS & ColumnLetter((x + 29)) & ",P3|" & ColumnLetter((x + 28)) & ",P3|"                'OtherTax
    XLS = XLS & ColumnLetter((x + 29)) & ",P3|" & ColumnLetter((x + 28)) & ",W8|"                'OtherTax

    XLS = XLS & ColumnLetter((x + 28)) & ",TC|" & ColumnLetter((x + 29)) & ",W8|" ' Total Tax"
    XLS = XLS & ColumnLetter((x + 30)) & ",P3|" & ColumnLetter((x + 30)) & ",cColor:3:0.0.0|" 'LTaxRate
    XLS = XLS & ColumnLetter((x + 30)) & ",TC|" & ColumnLetter((x + 30)) & ",W8|" ' LTaxRate

    Countx = DCount("AssetId", "SchDocsReportFinal")
    TitleX = ""
    TitleX = TitleX & DLookup("[Lease No]", "SchDocsReportFinal") & " | "
    TitleX = TitleX & "Sch: " & DLookup("SchNo", "SchDocsReportFinal") & " | "
    TitleX = TitleX & "Units: " & Countx & " "
    Call MessageUser("Update", "Making the 1st Tab", "...getting there...")
    Dim XLS2 As String
    If TwoTab = True Then GoTo MakeTwoTab

  '  Dim Count2 As Integer
  '  Count2 = DCount("Units", "ScheduleDocsReportExcelBase", "SchID=" & Schidx)
    XLS2 = "|D,BL|E,BC|E,BX|F,TC|G,TC|N,P4|O,TC|P,TC|"
    SendReport2Excel "Select * From ScheduleDocsReportLocations WHERE SchID=" & SchIDx & ";", "ATF_Locations", "ATF Location Detail", True, "" & XLS2 & "", 0, False, 1, False

    Call MessageUser("Update", "Making the 2nd Tab", "Four more tabs...")
    XLS2 = "|D,BL|E,BC|E,BX|I,$E|J,$E|K,$E|L,$E|M,$E|M,P4|O,$E|P,P4|R,TC|Y,P4|Z,TC|AH,$E|AE,ZZ|"
    SendReport2Excel "Select * From ScheduleDocsReportCostRentTax WHERE SchID=" & SchIDx & ";", "ATF_CostRentTax", "ATF Cost, Rent and Tax Detail", True, "" & XLS2 & "", 0, False, 2, False

    Call MessageUser("Update", "Making the 3rd Tab", "Three Tabs left...")
    XLS2 = "|D,BL|E,BC|E,BX|I,$E|J,$E|K,$E|M,$E|O,SR|P,SR|R,TC|W,P4|Y,P4|Z,P4|AA,P4|"
    SendReport2Excel "Select * From ScheduleDocsReportEconomics WHERE SchID=" & SchIDx & ";", "ATF_Economics", "ATF Economics Detail", True, "" & XLS2 & "", 0, False, 3, False

    Call MessageUser("Update", "Making the 4th Tab", "Down to Two Tabs...")
    XLS2 = "|A,BC|C,$E|D,$E|"
    SendReport2Excel "Select * From SchDocsVendorSummary;", "Vendor Summary", "Vendor Invoices", True, "" & XLS2 & "", 0, False, 4, False

    Call MessageUser("Update", "Making the 5th Tab", "One more to go...")
    XLS2 = "|A,BC|H,$E|I,$E|"
    SendReport2Excel "Select * From SchDocsVendorDetail;", "Vendor Invoices", "Vendor Invoices", True, "" & XLS2 & "", 0, False, 5, False

    Call MessageUser("Update", "Making the 6th Tab", "Excel will now open...")
    SendReport2Excel "Select * From SchDocsReportFinal", "Unit_Docs", "" & TitleX & "", True, "" & XLS & "", 0, False, 6, True
    'Countx
    Call MessageUser("Close")
    Exit Function
MakeTwoTab:
'****** added 2nd Tab for Sandra 2/17/2025
    Call MessageUser("Update", "Making the 2nd Tab", "Down to Two Tabs...")
    XLS2 = ""
    XLS2 = "|A,BC|C,$E|D,$E|"
    SendReport2Excel "Select * From SchDocsVendorSummary;", "Vendor Summary", "Vendor Invoices", True, "" & XLS2 & "", 0, False, 1, False

    Call MessageUser("Update", "Making the 1st Tab", "Excel will now open...")
    XLS2 = ""
    XLS2 = "M,TC|S,TC|W,TC|AB,TC|"
    XLS = XLS & XLS2
    SendReport2Excel "Select * From SchDocsReportFinal", "Unit_Docs", "" & TitleX & "", True, "" & XLS & "", 0, False, 2, True
    'Countx
    Call MessageUser("Close")
    Exit Function
FixMe:
MakeScheduleReport_Error:
    Call MessageUser("Close")
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MakeScheduleReport of Module CustomFunctions"
End Function
Function ColumnLetter(ColumnNumber As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = s
End Function

Function SplitMe(StingTSplit As String, StrDelimiter As String, Section As Integer)
    SplitMe = Split(StingTSplit, StrDelimiter)(Section)
End Function

Private Function rownumbers(Countx As Integer)
    Dim i As Integer
    For i = 1 To Countx
        SQLz = "Update Rmkt Set InvoiceRequestNum = " & i & " Where NOT ISNULL(Min(InvoiceRequestDate))"
     Next i
    End Function
Public Function ECount(Expr As String, Domain As String, Optional Criteria As String, Optional bCountDistinct As Boolean) As Variant
    On Error GoTo Err_Handler
    'Purpose:   Enhanced DCount() function, with the ability to count distinct.
    'Return:    Number of records. Null on error.
    'Arguments: Expr           = name of the field to count. Use square brackets if the name contains a space.
    '           Domain         = name of the table or query.
    '           Criteria       = any restrictions. Can omit.
    '           bCountDistinct = True to return the number of distinct values in the field. Omit for normal count.
    'Notes:     Nulls are excluded (whether distinct count or not.)
    '           Use "*" for Expr if you want to count the nulls too.
    '           You cannot use "*" if bCountDistinct is True.
    'Examples:  Number of customers who have a region: ECount("Region", "Customers")
    '           Number of customers who have no region: ECount("*", "Customers", "Region Is Null")
    '           Number of distinct regions: ECount("Region", "Customers", ,True)
    Dim db As DAO.DataBase
    Dim rs As DAO.Recordset
    Dim strSQL As String

    'Initialize to return Null on error.
    ECount = Null
    Set db = DBEngine(0)(0)

    If bCountDistinct Then
        'Count distinct values.
        If Expr <> "*" Then             'Cannot count distinct with the wildcard.
            strSQL = "SELECT " & Expr & " FROM " & Domain & " WHERE (" & Expr & " Is Not Null)"
            If Criteria <> vbNullString Then
                strSQL = strSQL & " AND (" & Criteria & ")"
            End If
            strSQL = strSQL & " GROUP BY " & Expr & ";"
            Set rs = db.OpenRecordset(strSQL)
            If rs.RecordCount > 0& Then
                rs.MoveLast
            End If
            ECount = rs.RecordCount     'Return the number of distinct records.
            rs.close
        End If
    Else
        'Normal count.
        strSQL = "SELECT Count(" & Expr & ") AS TheCount FROM " & Domain
        If Criteria <> vbNullString Then
            strSQL = strSQL & " WHERE " & Criteria
        End If
        Set rs = db.OpenRecordset(strSQL)
        If rs.RecordCount > 0& Then
            ECount = rs!TheCount        'Return the count.
        End If
        rs.close
    End If

Exit_Handler:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbExclamation, "ECount Error " & Err.Number
    Resume Exit_Handler
End Function
Public Function ELookup(Expr As String, Domain As String, Optional Criteria As Variant, _
                        Optional OrderClause As Variant) As Variant
    On Error GoTo Err_ELookup
    'Purpose:   Faster and more flexible replacement for DLookup()
    'Arguments: Same as DLookup, with additional Order By option.
    'Return:    Value of the Expr if found, else Null.
    '           Delimited list for multi-value field.
    'Author:    Allen Browne. allen@allenbrowne.com
    'Updated:   December 2006, to handle multi-value fields (Access 2007 and later.)
    'Examples:
    '           1. To find the last value, include DESC in the OrderClause, e.g.:
    '               ELookup("[Surname] & [FirstName]", "tblClient", , "ClientID DESC")
    '           2. To find the lowest non-null value of a field, use the Criteria, e.g.:
    '               ELookup("ClientID", "tblClient", "Surname Is Not Null" , "Surname")
    'Note:      Requires a reference to the DAO library.
    Dim db As DAO.DataBase      'This database.
    Dim rs As DAO.Recordset     'To retrieve the value to find.
    Dim rsMVF As DAO.Recordset     'Child recordset to use for multi-value fields.
    Dim varResult As Variant        'Return value for function.
    Dim strSQL As String            'SQL statement.
    Dim strOut As String            'Output string to build up (multi-value field.)
    Dim lngLen As Long              'Length of string.
    Const strcSep = ","             'Separator between items in multi-value list.

    'Initialize to null.
    varResult = Null

    'Build the SQL string.
    strSQL = "SELECT TOP 1 " & Expr & " FROM " & Domain
    If Not IsMissing(Criteria) Then
        strSQL = strSQL & " WHERE " & Criteria
    End If
    If Not IsMissing(OrderClause) Then
        strSQL = strSQL & " ORDER BY " & OrderClause
    End If
    strSQL = strSQL & ";"

    'Lookup the value.
    Set db = DBEngine(0)(0)
    Set rs = db.OpenRecordset(strSQL, dbOpenForwardOnly)
    If rs.RecordCount > 0 Then
        'Will be an object if multi-value field.
        If VarType(rs(0)) = vbObject Then
            Set rsMVF = rs(0).Value
            Do While Not rsMVF.EOF
                If rs(0).Type = 101 Then        'dbAttachment
                    strOut = strOut & rsMVF!fileName & strcSep
                Else
                    strOut = strOut & rsMVF![Value].Value & strcSep
                End If
                rsMVF.MoveNext
            Loop
            'Remove trailing separator.
            lngLen = Len(strOut) - Len(strcSep)
            If lngLen > 0& Then
                varResult = Left(strOut, lngLen)
            End If
            Set rsMVF = Nothing
        Else
            'Not a multi-value field: just return the value.
            varResult = rs(0)
        End If
    End If
    rs.close

    'Assign the return value.
    ELookup = varResult

Exit_ELookup:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Err_ELookup:
    MsgBox Err.Description, vbExclamation, "ELookup Error " & Err.Number
    Resume Exit_ELookup
End Function

Function ListBoxClear(lst As ListBox) As Boolean
    On Error GoTo Err_ClearList
    Dim varitem As Variant
    If lst.MultiSelect = 0 Then
        lst = Null
    Else
        For Each varitem In lst.ItemsSelected
            lst.Selected(varitem) = False
        Next
    End If
    ListBoxClear = True
Exit_ClearList:
    Exit Function
Err_ClearList:
    MsgBox Err.Description, vbExclamation, "ListBoxClear Error " & Err.Number
    Resume Exit_ClearList
End Function

Public Function ListBoxSelectAll(lst As ListBox) As Boolean
    On Error GoTo Err_Handler
    Dim lngRow As Long
    If lst.MultiSelect Then
        For lngRow = 0 To lst.ListCount - 1
            lst.Selected(lngRow) = True
        Next
        ListBoxSelectAll = True
    End If
Exit_Handler:
    Exit Function
Err_Handler:
    MsgBox Err.Description, vbExclamation, "ListBoxSelectAll Error " & Err.Number
    Resume Exit_Handler
End Function

Function GetNextWorkDay(StartDate As Date, DaysInFuture As Integer)
    Dim TempDate As Date
    TempDate = StartDate + DaysInFuture
    If Weekday(TempDate) = 7 Then TempDate = TempDate + 2
    If Weekday(TempDate) = 1 Then TempDate = TempDate + 1
    GetNextWorkDay = TempDate
End Function
Function GetNextWorkDayInDays(StartDate As Date, DaysInFuture As Integer)
    Dim TempDate As Date
    TempDate = StartDate + DaysInFuture
    If Weekday(TempDate) = 7 Then TempDate = TempDate + 2
    If Weekday(TempDate) = 1 Then TempDate = TempDate + 1
    GetNextWorkDayInDays = DateDiff("d", Date, TempDate)
End Function
Function MakeScheduleWorkTable(SchIDx As Long, Optional GroupIDx As Long)
    DoCmd.SetWarnings False
    SQLz = "Delete * From Schedule_Group_WorkTable;"
    DoCmd.RunSQL (SQLz)
    SQLz = "INSERT INTO Schedule_Group_WorkTable ( ScheduleId, GroupId ) SELECT " & SchIDx & ", " & GroupIDx & ";"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Function
Function MakeCustodianSubmission(xSchiD As Integer)
    Dim Countx As Integer
    Countx = DCount("Lessee", "vw_CustodianTitleTrust2")
    If Countx < 1 Then
        MsgBox "There is missing data this is preventing this report from running," & vbNewLine & "(Such as a missing Bank or Portfolio #).", vbCritical, "No records"
        Exit Function
    End If
   ' Call UpdateUnitTitleNum
   ' Call UpdateUnitTitleState
    Dim DateX, TrackX As String
    Dim Response As String
    DateX = Nz(DLookup("DateSent", "vw_CustodianTitleTrust2"), Date)
    Response = MsgBox("Do you want to send an Custodian Submission Email?", vbQuestion + vbYesNo + vbCritical, "Email it?")
    If Response = vbYes Then
        ExcelCustodian2 "CustodianTitleTrust2DataFile", "C:\Faas\" & FaUserName & "\Temp\CustodianTitleTrust2DataFile.xlsx", "DataFile", "Close", , Countx, "Data Files - Sent " & DateX, False, "CC"
        ExcelCustodian1 "CustodianTitleTrust2", "C:\Faas\" & FaUserName & "\Temp\CustodianTitleTrust2.xlsx", "Custodian", "Close", , Countx, "Equipment Files - Sent " & DateX, False, "CC"
        Call EmailToTrust(xSchiD)
    Else
        ExcelCustodian2 "CustodianTitleTrust2DataFile", "C:\FAAS\" & FaUserName & "\temp\CustodianTitleTrust2DataFile.xlsx", "DataFile", , , Countx, "Data Files - Sent " & DateX, False, "CC"
        ExcelCustodian1 "CustodianTitleTrust2", "C:\Faas\" & FaUserName & "\Temp\CustodianTitleTrust2.xlsx", "Custodian", , , Countx, "Equipment Files - Sent " & DateX, False, "CC"
    End If
    Response = MsgBox("Do you want to log these " & Countx & " Titles as 'Sent To Custodian'?", vbQuestion + vbYesNo + vbCritical, "Log it?")
    If Response = vbYes Then
        Dim DateSent As String
AskDate:
        DateSent = InputBox("Provide the 'Sent Date' for these titles." & vbNewLine & _
                          "          Please enter that date, below." & vbNewLine & _
                          "          Today's date is the default." & vbNewLine & vbNewLine & _
                          "          ( You must use a date )" _
                 , "Custodian Titles, Sent Date", Date)
        If Not IsDate(DateSent) Then GoTo AskDate
        SQLz = "INSERT INTO UnitTitleLog ( UnitRef, UnitTitleLogDate, EntryTypeId, EntryReasonId, TitleSt, TitleNumber, EntryReason, UserID )" & _
               "SELECT Units.UnitID, #" & [DateSent] & "# AS SentDate, 2 AS Type, 4 AS Reasonid, vw_CustodianTitleTrust2.[State of Title], " & _
               "vw_CustodianTitleTrust2.TitleNum, 'Sent To Custodian' AS Reason, '" & FaUserName & "' " & _
               "FROM vw_CustodianTitleTrust2 INNER JOIN Units ON vw_CustodianTitleTrust2.VIN = Units.UnitVIN " & _
               "ORDER BY Units.AssetID_FATitleLog;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.RunSQL ("Delete From TitleLogStateTemp")
        DoCmd.RunSQL ("INSERT INTO TitleLogStateTemp ( UnitID, TitleSt ) " & _
                      "SELECT Units.UnitID, TL.TitleSt FROM Units INNER JOIN UnitTitleHistoryLastRecv As TL ON Units.UnitID = TL.UnitRef " & _
                      "WHERE TL.TitleSt Is Not Null And TL.TitleSt <> ' ' And Units.UnitTitleState Is Null;")
        DoCmd.RunSQL ("UPDATE Units INNER JOIN TitleLogStateTemp ON Units.UnitID = TitleLogStateTemp.UnitID SET Units.UnitTitleState = [TitleLogStateTemp].[TitleSt];")
        DoCmd.RunSQL ("Delete From TitleLogStateTemp")
        DoCmd.SetWarnings True
        MsgBox "Title Log Updated For " & Countx & " Units", vbInformation, "Log Updated"
    End If
  '  Call UpdateUnitTitleNum(Me.UnitID)
  '  Call UpdateUnitTitleState(Me.UnitID)
End Function
Function UpdateUnitTitleState(UnitIDx As Long)

    Dim CurrentTitleState, ThisTitleState As String
    CurrentTitleState = Nz(DLookup("UnitTitleState", "Units", "UnitID=" & UnitIDx), "")
    ThisTitleState = Nz(DLookup("TitleSt", "UnitTitleHistoryLastRecv", "UnitRef=" & UnitIDx), "")
    If CurrentTitleState = ThisTitleState Then Exit Function
    If ThisTitleState = "" Then Exit Function
    If CurrentTitleState <> ThisTitleState Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("UPDATE Units SET Units.UnitTitleState = '" & ThisTitleState & "' Where UnitID=" & UnitIDx)
        DoCmd.SetWarnings True
    End If
    'Dim RST    As DAO.Recordset
    'Set RST = CurrentDb.OpenRecordset( _
     '          "SELECT Units.UnitID, TL.TitleSt FROM Units INNER JOIN UnitTitleHistoryLastRecv AS TL ON Units.UnitID = TL.UnitRef " & _
     '          "WHERE TL.TitleSt <> Nz([Units].[UnitTitleState],'');")

    'If RST.RecordCount > 0 Then
    '    DoCmd.RunSQL ("Delete From TitleLogStateTemp")
    '    DoCmd.RunSQL ("INSERT INTO TitleLogStateTemp ( UnitID, TitleSt ) " & _
         '                  "SELECT Units.UnitID, TL.TitleSt FROM Units INNER JOIN UnitTitleHistoryLastRecv As TL ON Units.UnitID = TL.UnitRef " & _
         "WHERE TL.TitleSt Is Not Null And TL.TitleSt <> ' ' AND Nz(Units.UnitTitleState,'') <> TL.TitleSt;")
    '    DoCmd.RunSQL ("UPDATE Units INNER JOIN TitleLogStateTemp ON Units.UnitID = TitleLogStateTemp.UnitID SET Units.UnitTitleState = [TitleLogStateTemp].[TitleSt];")

    ' DoCmd.RunSQL ("Delete From TitleLogStateTemp")
    ' DoCmd.RunSQL ("INSERT INTO TitleLogStateTemp ( UnitID, TitleSt ) " & _
      "SELECT Distinct TL.UnitRef, TL.TitleSt FROM UnitTitleHistoryLastRecv As TL INNER JOIN UnitTitleLog ON " & _
      "(TL.TitleNumber = UnitTitleLog.TitleNumber) AND (TL.UnitRef = UnitTitleLog.UnitRef) " & _
      "WHERE TL.TitleSt Is Not Null AND Nz(UnitTitleLog.TitleSt,'')='' ;")

    'DoCmd.RunSQL ("UPDATE TitleLogStateTemp INNER JOIN UnitTitleLog ON TitleLogStateTemp.UnitID = UnitTitleLog.UnitRef " & _
     "Set UnitTitleLog.TitleSt = [TitleLogStateTemp].[TitleSt] " & _
     "Where UnitTitleLog.TitleSt is null;")
    'DoCmd.RunSQL ("Delete From TitleLogStateTemp")

End Function
Function UpdateUnitTitleNum(UnitIDx As Long)

    Dim CurrentTitleNum, ThisTitleNum As String
    CurrentTitleNum = Nz(DLookup("UnitTitleNumber", "Units", "UnitID=" & UnitIDx), "")
    ThisTitleNum = Nz(DLookup("TitleNumber", "UnitTitleHistoryLastRecv", "UnitRef=" & UnitIDx), "")
    If CurrentTitleNum = ThisTitleNum Then Exit Function
    If ThisTitleNum = "" Then Exit Function
    If CurrentTitleNum <> ThisTitleNum Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL ("UPDATE Units SET Units.UnitTitleNumber = '" & ThisTitleNum & "' Where UnitID=" & UnitIDx)
        DoCmd.SetWarnings True
    End If
    'Dim RST    As DAO.Recordset
    'Set RST = CurrentDb.OpenRecordset( _
     "SELECT Units.UnitID, TL.TitleNumber FROM Units INNER JOIN UnitTitleHistoryLastRecv AS TL ON Units.UnitID = TL.UnitRef " & _
     "WHERE TL.TitleNumber <> Nz([Units].[UnitTitleNumber],'');")
    'DoCmd.SetWarnings False
    'If RST.RecordCount > 0 Then
    '    DoCmd.RunSQL ("Delete From TitleLogNumTemp")
    '    DoCmd.RunSQL ("INSERT INTO TitleLogNumTemp ( UnitID, TitleNumber ) " & _
         "SELECT Units.UnitID, TL.TitleNumber FROM Units INNER JOIN UnitTitleHistoryLastRecv As TL ON Units.UnitID = TL.UnitRef " & _
         "WHERE TL.TitleNumber Is Not Null And TL.TitleNumber <>' ' AND Nz(Units.UnitTitleNumber,'') <> TL.TitleNumber;")
    '    DoCmd.RunSQL ("UPDATE Units INNER JOIN TitleLogNumTemp ON Units.UnitID = TitleLogNumTemp.UnitID SET Units.UnitTitleNumber = [TitleLogNumTemp].[TitleNumber];")
    'End If

    'DoCmd.RunSQL ("Delete From TitleLogNumTemp")
    'DoCmd.RunSQL ("INSERT INTO titlelognumtemp ( UnitID, TitleNumber ) " & _
     "SELECT TL.UnitRef, TL.TitleNumber FROM UnitTitleHistoryLastRecv AS TL INNER JOIN UnitTitleLog ON TL.UnitRef = UnitTitleLog.UnitRef " & _
     "WHERE TL.TitleNumber Is Not Null AND UnitTitleLog.EntryReasonId<>18 AND Nz(UnitTitleLog.TitleNumber,'') ='';")

    'DoCmd.RunSQL ("UPDATE TitleLogNumTemp INNER JOIN UnitTitleLog ON TitleLogNumTemp.UnitID = UnitTitleLog.UnitRef " & _
     "Set UnitTitleLog.TitleNumber = [TitleLogNumTemp].[TitleNumber] " & _
     "Where Nz(UnitTitleLog.TitleNumber,'') <>  [TitleLogNumTemp].[TitleNumber] AND UnitTitleLog.EntryReasonId<>18;")
    'DoCmd.RunSQL ("Delete From TitleLogNumTemp")
    'DoCmd.SetWarnings True
    'RST.Close
    'Set RST = Nothing
End Function
Function MakeClientUnitSummary(ClientIDx As Long)
    Dim XLS As String
    Dim Countx As Integer
    SQLz = "Select * From Client_LPAF Where ClientId=" & ClientIDx & ";"
    Countx = DCount("clientID", "Client_LPAF", "ClientId=" & ClientIDx)
    If Nz(Countx, 0) <= 0 Then
        MsgBox "No Matching Units", vbInformation, "No Result"
        Exit Function
    End If
    XLS = "|A,TC|B,TC|G,$E|H,$E|J,TC|M,$E|N,TC|O,$E|P,TC|Q,$E|R,$E|S,HH|T,HH|U,HH|"
    SendReport2Excel SQLz, "Report", "Report for: " & DLookup("ClientShortName", "vw_SixKeys", "ClientId=" & ClientIDx) & " All Schedules - Units: " & Countx, True, XLS, Countx
End Function

Public Function AskToDoTask(TaskX As String) As Boolean
    AskToDoTask = True
    Dim Askit As String
    Askit = MsgBox(TaskX & vbNewLine & "Do you agree?", vbQuestion + vbYesNo + vbCritical, AtlaasVersion & " Do Task?")
    If Askit <> 6 Then AskToDoTask = False

End Function
'---------------------------------------------------------------------------------------
' Procedure : ListOpenFrms
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Returns a list of all the loaded forms (preview or design)
'             separated by ; (ie: Form1;Form2;Form3)
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites.
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2009-Oct-31                 Initial Release
'---------------------------------------------------------------------------------------
Public Function ListOpenFrms()
    On Error GoTo Error_Handler
    Dim DbF As Form
    Dim DbO As Object
    Dim Frms As Variant
    Set DbO = Application.Forms    'Collection of all the open forms
    For Each DbF In DbO    'Loop all the forms
        Frms = Frms & ";" & DbF.Name
    Next DbF
    If Len(Frms) > 0 Then
        Frms = Right(Frms, Len(Frms) - 1)   'Truncate initial ;
    End If
    ListOpenFrms = Frms
    Exit Function
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
           Err.Number & vbCrLf & "Error Source: ListOpenFrms" & vbCrLf & "Error Description: " & _
           Err.Description, vbCritical, "An Error has Occured!"
    Exit Function
End Function
Public Function HSpaces(NumOfSpaces As Integer) As String
'TopLeftCorner = ChrW(9484)
'TopRghtCorner = ChrW(9488)
'LineSpaces = ChrW(9472)
    Dim ii As Integer
    For ii = 1 To NumOfSpaces
        HSpaces = HSpaces & "&nbsp;"
    Next ii
End Function
Public Function LineSpace(NumOfSpaces As Integer) As String
'TopLeftCorner = ChrW(9484)
'TopRghtCorner = ChrW(9488)
'LineSpaces = ChrW(9472)
    Dim ii As Integer
    For ii = 1 To NumOfSpaces
        LineSpace = LineSpace & ChrW(9472)
    Next ii
End Function
Public Function TLeftLine() As String
    TLeftLine = ChrW(9484)
End Function
Public Function TRightLine() As String
    TRightLine = ChrW(9488)
End Function
Public Function Coalesce(ParamArray varValues()) As Variant
'returns the first non null value, similar to SQL Server Coalesce() function
'Patrick Honorez --- www.idevlop.com
    Dim i As Integer
    Coalesce = Null
    For i = LBound(varValues) To UBound(varValues)
        If Not IsNull(varValues(i)) Then
            Coalesce = varValues(i)
            Exit Function
        End If
    Next
End Function

Public Function ValidateField(FieldNamex As String, TableNamex As String, Optional AssetIDx As Long, Optional GroupIDx As Long, Optional ScheduleIDx As Long) As String

    Dim WhereField, GroupDetailx As String
    Dim ControlField As String
    Dim WhereValue As Long
    Dim x, Cntx, Missingi, Seti As Integer
    On Error GoTo ValidateField_Error
    GroupDetailx = ""
    If AssetIDx > 0 Then
        WhereValue = AssetIDx
        WhereField = "UnitID"
        ControlField = "UnitID"
    End If
    If GroupIDx > 0 Then
        WhereValue = GroupIDx
        WhereField = "SGrpID"
        ControlField = "SGrpID"
        If TableNamex = "Units" Then WhereField = "GroupRef"
        If TableNamex = "Units" Then ControlField = "UnitID"
    End If
    If ScheduleIDx > 0 Then
        WhereField = "SchID"
        WhereValue = ScheduleIDx
        ControlField = "SchID"
        If TableNamex = "SchGrp" Then WhereField = "SgrpSchID"
        If TableNamex = "SchGrp" Then ControlField = "SGrpID"
    End If

    If ScheduleIDx > 0 And TableNamex = "Units" Then
        Dim RST As DAO.Recordset
        Dim SQLz, GroupList As String
        WhereField = "GroupRef"
        ControlField = "UnitID"
        GroupList = ""
        SQLz = "Select * From SchGrp Where SgrpSchID=" & ScheduleIDx & ";"
        Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
        RST.MoveLast
        RST.MoveFirst
        If RST.RecordCount < 1 Then
            RST.close
            Set RST = Nothing
            Seti = 0
            Missingi = 1
        Else
            RST.MoveFirst
            Seti = 0
            Cntx = 0
            Missingi = 0
            For x = 1 To RST.RecordCount
                If x = 1 Then GroupList = RST!SGrpID Else GroupList = GroupList & ", " & RST!SGrpID
                '         Seti = Seti + DCount("UnitID", TableNamex, "GroupRef=" & RST!SGrpID & " And Not IsNull(" & FieldNamex & ")")
                Missingi = Missingi + DCount("UnitID", TableNamex, "GroupRef=" & RST!SGrpID & " And Nz(" & FieldNamex & ",0)=0")
                Cntx = DCount("UnitID", TableNamex, "GroupRef=" & RST!SGrpID & " And Nz(" & FieldNamex & ",0)=0")
                If FieldNamex = "UnitSynResid" And DCount("UnitID", "Units", "UnitSubType='Auto Hauler' AND GroupRef=" & RST!SGrpID & "") > 0 Then Cntx = 0
                If FieldNamex = "UnitSalesResid" And DCount("UnitID", "Units", "UnitSubType='Auto Hauler' AND GroupRef=" & RST!SGrpID & "") > 0 Then Cntx = 0
                If Cntx > 0 Then GroupDetailx = GroupDetailx & vbNewLine & "   " & Cntx & " units are missing data in Group: " & DLookup("UnitGroup", "vw_SixKeys", "sGrpID=" & RST!SGrpID)
                RST.MoveNext
            Next x
        End If
        ' If FieldNamex = "UnitSynResid" And EqpType = "Trailer" Then Missingi = 0
    Else
        Dim EqpType As String
        EqpType = Nz(DLookup("SgrpEq1", "SchGrp", "SGrpID=" & GroupIDx & ""), "")
        '  Seti = DCount(FieldNamex, TableNamex, WhereField & "=" & WhereValue & "")
        Missingi = DCount(ControlField, TableNamex, WhereField & "=" & WhereValue & " And Nz(" & FieldNamex & ",0)=0")
        If Missingi > 0 Then GroupDetailx = GroupDetailx & vbNewLine & "   " & " is missing in Group: " & DLookup("UnitGroup", "vw_SixKeys", "sGrpID=" & GroupIDx)
        If FieldNamex = "SGprAllow1" And EqpType = "Trailer" Then Missingi = 0
    End If
    ValidateField = ""
    'ValidateField = "# records Set:" & Seti
    'If Missingi > 0 Then ValidateField = ValidateField & "Number of records missing data: " & Missingi
    If Missingi > 0 Then ValidateField = ValidateField & GroupDetailx
    On Error GoTo 0
    Exit Function
ValidateField_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidateField of Module CustomFunctions"
End Function
Public Function SchHealthCheck(FieldNamex As String, TableNamex As String, Optional AssetIDx As Long, Optional GroupIDx As Long, Optional ScheduleIDx As Long) As String
    Dim WhereField, GroupDetailx As String
    Dim ControlField As String
    Dim WhereValue As Long
    Dim x, Cntx, Missingi, Seti As Integer
    On Error GoTo SchHealthCheck_Error
    GroupDetailx = ""
    If AssetIDx > 0 Then
        WhereValue = AssetIDx
        WhereField = "UnitID"
        ControlField = "UnitID"
    End If
    If GroupIDx > 0 Then
        WhereValue = GroupIDx
        WhereField = "SGrpID"
        ControlField = "SGrpID"
        If TableNamex = "Units" Then WhereField = "GroupRef"
        If TableNamex = "Units" Then ControlField = "UnitID"
    End If
    If ScheduleIDx > 0 Then
        WhereField = "SchID"
        WhereValue = ScheduleIDx
        ControlField = "SchID"
        If TableNamex = "SchGrp" Then WhereField = "SgrpSchID"
        If TableNamex = "SchGrp" Then ControlField = "SGrpID"
    End If

    If ScheduleIDx > 0 And TableNamex = "Units" Then
        Dim RST As DAO.Recordset
        Dim SQLz, GroupList As String
        WhereField = "GroupRef"
        ControlField = "UnitID"
        GroupList = ""
        SQLz = "Select SGrpID From SchGrp Where SgrpSchID=" & ScheduleIDx & ";"
        Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
        RST.MoveLast
        RST.MoveFirst
        If RST.RecordCount < 1 Then
            RST.close
            Set RST = Nothing
            Seti = 0
            Missingi = 1
        Else
            RST.MoveFirst
            Seti = 0
            Cntx = 0
            Missingi = 0
            For x = 1 To RST.RecordCount
                If x = 1 Then GroupList = RST!SGrpID Else GroupList = GroupList & ", " & RST!SGrpID
                '         Seti = Seti + DCount("UnitID", TableNamex, "GroupRef=" & RST!SGrpID & " And Not IsNull(" & FieldNamex & ")")
                Missingi = Missingi + DCount("UnitID", TableNamex, "GroupRef=" & RST!SGrpID & " And Nz(" & FieldNamex & ",0)=0")
                Cntx = DCount("UnitID", TableNamex, "GroupRef=" & RST!SGrpID & " And Nz(" & FieldNamex & ",0)=0")
                If Cntx > 0 Then GroupDetailx = GroupDetailx & "   " & Cntx & " units are missing data in Group: " & DLookup("UnitGroup", "vw_SixKeys", "sGrpID=" & RST!SGrpID) & vbNewLine
                RST.MoveNext
            Next x
        End If
    Else
        '  Seti = DCount(FieldNamex, TableNamex, WhereField & "=" & WhereValue & "")
        Missingi = DCount(ControlField, TableNamex, WhereField & "=" & WhereValue & " And Nz(" & FieldNamex & ",0)=0")
        If Missingi > 0 Then GroupDetailx = GroupDetailx & "" & " is missing in Group: " & DLookup("UnitGroup", "vw_SixKeys", "sGrpID=" & GroupIDx) & vbNewLine
    End If
    SchHealthCheck = ""
    'SchHealthCheck = "# records Set:" & Seti
    'If Missingi > 0 Then SchHealthCheck = SchHealthCheck & "Number of records missing data: " & Missingi
    If Missingi > 0 Then SchHealthCheck = SchHealthCheck & GroupDetailx
    On Error GoTo 0
    Exit Function
SchHealthCheck_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SchHealthCheck of Module CustomFunctions"
End Function

Public Function MakeHealthReport()
    Dim RST As DAO.Recordset
    Dim SQLz, GroupList As String
    Dim x As Integer
    Dim PDRentX, SYNDRentX, Assigned_IndicatorX, SchOut, SchIn As String
    SQLz = "Delete * From ScheduleHealthCheckWorkTable"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    SQLz = "Select * From ScheduleHealthCheck Order By Client, SchNo"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then
        RST.close
        Set RST = Nothing
    Else
        RST.MoveFirst
        For x = 1 To RST.RecordCount
            SQLz = "INSERT INTO ScheduleHealthCheckWorkTable ( ClientID, SchID, GroupID, SchNo, Client, SynRent, PDRent )"
            SQLz = SQLz & " SELECT "
            SQLz = SQLz & RST!ClientID
            SQLz = SQLz & RST!SchID
            SQLz = SQLz & RST!GroupID
            SQLz = SQLz & RST!SchNo
            SQLz = SQLz & RST!Client
            SQLz = SQLz & "," & SchHealthCheck("UNITRENT", "Units", , RST!GroupID, RST!SchID)
            SQLz = SQLz & "," & SchHealthCheck("PDRent", "Units", , RST!GroupID, RST!SchID)
            '  DoCmd.SetWarnings False
            DoCmd.RunSQL (SQLz)
            DoCmd.SetWarnings True
            RST.MoveNext
        Next x
    End If
End Function
Public Function UnitsExtends(AssetIDx As Long) As String
    UnitsExtends = Nz(ConcatRelated("ExtendDesc", "vw_UnitExtends", "UnitID=" & AssetIDx & "", "StartD", " "), "")
End Function
Public Function BLCDInSchedule(SchIDx As Long) As String
    BLCDInSchedule = Nz(ConcatRelated("GBLCD", "vw_ScheduleDates", "SchId=" & SchIDx, "GBLCD", "" & vbNewLine & ""), "")
End Function
Public Function AssignmentDateInSchedule(SchIDx As Long) As String
    AssignmentDateInSchedule = Nz(ConcatRelated("GAssignmentDate", "vw_ScheduleDates", "SchId=" & SchIDx, "GBLCD", "" & vbNewLine & ""), "")
End Function
Public Function FundDateInSchedule(SchIDx As Long) As String
    FundDateInSchedule = Nz(ConcatRelated("GFundDate", "vw_ScheduleDates", "SchId=" & SchIDx, "GBLCD", "" & vbNewLine & ""), "")
End Function
Public Function UnitDescInGroup2(GroupIDx As Long) As String
    UnitDescInGroup2 = ConcatRelated("ShortDesc", "TFRDesc", "GroupID=" & GroupIDx, "ShortDesc", "" & vbNewLine & "")
End Function
Public Function RALOutSchedule(SchIDx As Long) As String
    'RALOutSchedule = ConcatRelated("GRALOut", "vw_ScheduleDates", "SchId=" & SchIDx, "GBLCD" "" & vbNewLine & "")
End Function
Public Function RALBackSchedule(SchIDx As Long) As String
   '- RALBackSchedule = ConcatRelated("GRALBack", "vw_ScheduleDates", "SchId=" & SchIDx, "GBLCD", "" & vbNewLine & "")
End Function
Public Function TMTSigned(SchIDx As Long) As String
    TMTSigned = Nz(ConcatRelated("TMTSigned", "vw_ScheduleDates", "SchId=" & SchIDx, "GBLCD", "" & vbNewLine & ""), "n/a")
End Function
Public Function GroupsNoinSchedule(SchIDx As Long, Optional GroupIDx As Long, Optional RichText As Boolean, Optional FundedOnly As Boolean) As String
    Dim TempStr1, TempStr2 As String
    If GroupIDx > 0 Then
       If FundedOnly = False Then GroupsNoinSchedule = ConcatRelated("sGrpNo", "schGrp", "sgrpID=" & GroupIDx, "sGrpNo", ", ")
       If FundedOnly = True Then GroupsNoinSchedule = ConcatRelated("sGrpNo", "schGrp", "Assigned_Indicator='F' AND sgrpID=" & GroupIDx, "sGrpNo", ", ")
    Else
       If FundedOnly = False Then GroupsNoinSchedule = ConcatRelated("sGrpNo", "schGrp", "sgrpSchID=" & SchIDx, "sGrpNo", ", ")
       If FundedOnly = True Then GroupsNoinSchedule = ConcatRelated("sGrpNo", "schGrp", "Assigned_Indicator='F' AND sgrpSchID=" & SchIDx, "sGrpNo", ", ")
    End If
    If Len(GroupsNoinSchedule) > 6 Then
        GroupsNoinSchedule = Replace(GroupsNoinSchedule, ", , ", ", ")
        TempStr1 = Left(GroupsNoinSchedule, InStrRev(GroupsNoinSchedule, ",") - 1)
        '        TempStr2 = Right(GroupsNoinSchedule, Len(GroupsNoinSchedule) - InStrRev(GroupsNoinSchedule, ","))
        TempStr2 = Right(TempStr1, Len(TempStr1) - InStrRev(TempStr1, ","))
        TempStr1 = Left(TempStr1, InStrRev(TempStr1, ",") - 1)
       ' If Right(TempStr2, 2) = ", " Then TempStr2 = Left(TempStr2, Len(TempStr2) - 2)
        GroupsNoinSchedule = "Groups " & TempStr1 & " and " & TempStr2
        GroupsNoinSchedule = Replace(GroupsNoinSchedule, "  ", " ")
        If RichText = True Then
            GroupsNoinSchedule = Replace(GroupsNoinSchedule, "and", "</b>and<b>")
            GroupsNoinSchedule = Replace(GroupsNoinSchedule, "Groups", "Groups<b>")
            GroupsNoinSchedule = GroupsNoinSchedule & "</b>"
        End If
    Else
        GroupsNoinSchedule = "Group " & Replace(GroupsNoinSchedule, ", ", "")
    End If

End Function
Public Function ScheduleGroupsInLine(SchIDx As Long) As String
    Dim ScheduleX, TempStr1, TempStr2 As String
    ScheduleX = DLookup("Schedule", "vw_SixKeys", "SchID=" & SchIDx)

    If DCount("sGrpID", "schGrp", "sgrpSchID=" & SchIDx) > 1 Then
        TempStr1 = ConcatRelated("sGrpNo", "schGrp", "sgrpSchID=" & SchIDx, "sGrpNo", ") and (Group ")
        TempStr1 = "(Group " & TempStr1
        TempStr1 = Replace(TempStr1, "(Gro", "Schedule No. " & ScheduleX & " (Gro")
        TempStr2 = ConcatRelated("sGrpNo", "schGrp", "sgrpSchID=" & SchIDx, "sGrpNo", " and Group ")
        TempStr2 = "Group " & TempStr2
        ScheduleGroupsInLine = TempStr1 & ", and each " & TempStr2
    Else
        TempStr1 = "Schedule No. " & ScheduleX & " Group " & ConcatRelated("sGrpNo", "schGrp", "sgrpSchID=" & SchIDx, "sGrpNo", ")")
        TempStr2 = ConcatRelated("sGrpNo", "schGrp", "sgrpSchID=" & SchIDx, "sGrpNo", " Group ")
        TempStr2 = "Group " & TempStr2
        ScheduleGroupsInLine = TempStr1 & ", and " & TempStr2
    End If
End Function
Sub Tester()
    MsgBox CleanString("3d1fgd4g1dg5d9gdg")
End Sub

Public Function NumInString(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Global = True
        .Pattern = "[^\d]+"
        NumInString = .Replace(strIn, vbNullString)
    End With
    If Not IsNumeric(NumInString) Then NumInString = 0
End Function
Public Function UnitRentsinGroup(GroupIDx As Long, Optional NoLineBreaks As Boolean) As String
    Dim TempStr1, TempStr2, Suffix As String
    If NoLineBreaks = True Then Suffix = " " Else Suffix = vbNewLine
    If GroupIDx <> 0 Then
        UnitRentsinGroup = "<u>" & Format(DSum("UnitRent", "Units", "GroupRef=" & GroupIDx), "$#,#.##") & "</u>" & Suffix & "(" & ConcatRelated("RentLine", "RALRents", "GroupID=" & GroupIDx, "Units desc", "" & Suffix & "")
    Else
    End If
    UnitRentsinGroup = UnitRentsinGroup & ") per month,"
     If DLookup("LeasePaidAdvArr", "schgrp", "sGrpID=" & GroupIDx) = "ADV" Then
        UnitRentsinGroup = UnitRentsinGroup & " <b>in advance</b>"
     Else
        UnitRentsinGroup = UnitRentsinGroup & " <b>in arrears</b>"
     End If
End Function
Public Function UnitDescInGroup(GroupIDx As Long, Optional NoLineBreaks As Boolean, Optional NoCOAComment As Boolean) As String
    Dim TempStr1, TempStr2, Suffix As String
    If NoLineBreaks = True Then Suffix = " " Else Suffix = vbNewLine
    If GroupIDx <> 0 Then
        UnitDescInGroup = ConcatRelated("UnitDesc", "RALUnitDesc", "GroupID=" & GroupIDx, "Unitdesc", "" & Suffix & "")
    Else
    End If
    If NoCOAComment <> True Then UnitDescInGroup = UnitDescInGroup & vbNewLine & "as described on the Certificate(s) of Acceptance"
End Function
'DoCmd.OutputTo acOutputReport, "RALReport", acFormatRTF, "c:\FAAS\" & FAUserName & "\temp\Filepath.rtf"
