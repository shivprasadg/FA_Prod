Attribute VB_Name = "DocumentFunctions"
Option Compare Database
Option Explicit
Dim SQLz As String
Dim RST    As DAO.Recordset
Dim WordApp  As Word.Application
Dim WordDoc  As Word.Document
Dim WordTable As Word.Table


Public Sub MakePAI(iSchID As Long)

    Dim db As DAO.DataBase
    Dim rsHeader As DAO.Recordset
    Dim rsDetail As DAO.Recordset

    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim tbl As Word.Table
    Dim rng As Word.Range

    Dim sqlHeader As String
    Dim sqlDetail As String

    Dim faSchText As String
    Dim reText As String
    Dim grandTotal As Currency
    Dim rowIndex As Long
    Dim logoEnd As Long

    Set db = CurrentDb


    '-------------------------
    ' HEADER DATA (PAITemp1)
    '-------------------------
    sqlHeader = "SELECT * FROM PAIBase WHERE SchID=" & iSchID & ";"
    Set rsHeader = db.OpenRecordset(sqlHeader, dbOpenDynaset, dbSeeChanges)
    If rsHeader.EOF Then
        MsgBox "Sorry, No Data found for That Schedule", vbExclamation, "No Data, No Report"
        Exit Sub
    End If

    '-------------------------
    ' DETAIL DATA (PAIBase)
    '-------------------------
    sqlDetail = "SELECT Contract, GroupRent, ClientCompanyName, " & _
        "Schedule, GroupNo, UnitsInGroup, RentDueDates " & _
        "FROM PAIBase WHERE SchID=" & iSchID & " ORDER BY GroupNo;"
    Set rsDetail = db.OpenRecordset(sqlDetail, dbOpenDynaset)
    If rsDetail.EOF Then
        MsgBox "No detail rows found for SchID=" & iSchID, vbExclamation
        Exit Sub
    End If

    ' Compute Grand Total
    grandTotal = 0
    rsDetail.MoveFirst
    Do Until rsDetail.EOF
        grandTotal = grandTotal + Nz(rsDetail!GroupRent, 0)
        rsDetail.MoveNext
    Loop
    rsDetail.MoveFirst

    '-------------------------
    ' START WORD
    '-------------------------
    Set WordApp = CreateObject("Word.Application")
    Set WordDoc = WordApp.Documents.Add

    ' Narrow margins
    With WordDoc.PageSetup
        .TopMargin = 25
        .BottomMargin = 25
        .LeftMargin = 16
        .RightMargin = 16
    End With

    '-------------------------
    ' INSERT LOGO (SAFE, NOT OVERWRITTEN)
    '-------------------------
    Dim logoPath As String
    logoPath = "C:\FAAS\Images\Smalllogo.gif"

    If Dir(logoPath) = "" Then
        MsgBox "Logo not found: " & logoPath, vbCritical
        Exit Sub
    End If

    Dim shp As Word.InlineShape
    Set shp = WordDoc.Range(0, 0).InlineShapes.AddPicture( _
        fileName:=logoPath, _
        LinkToFile:=False, _
        SaveWithDocument:=True)

    ' Resize logo
    With shp
        .LockAspectRatio = True
        .Width = 160
    End With
    logoEnd = shp.Range.End

    ' Add spacing after logo
    WordDoc.Range(logoEnd, logoEnd).Text = vbTab & vbTab
    '   WordDoc.Range(shp.Range.End, shp.Range.End).InsertParagraphAfter

    '-------------------------
    ' TOP HEADER TEXT
    '-------------------------
    Set rng = WordDoc.Range(logoEnd + 1, logoEnd + 1)
    rng.Text = "Paying Instructions Under Paying Agency Agreement"
    rng.Font.Bold = True
    rng.Font.Size = 12

    rng.Collapse wdCollapseEnd
    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd

    rng.Text = "Assignee Wire Instructions:"
    rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
    rng.Font.Bold = True
    rng.Font.Size = 11

    rng.Collapse wdCollapseEnd
    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd

    '-------------------------
    ' TABLE 1 — WIRE INSTRUCTIONS
    '-------------------------
    Set tbl = WordDoc.tables.Add(rng, 6, 2)
    tbl.Style = "Table Grid"
    tbl.rows.Alignment = wdAlignRowLeft

    tbl.Columns(1).PreferredWidthType = wdPreferredWidthPercent
    tbl.Columns(1).PreferredWidth = 15
    tbl.Columns(2).PreferredWidthType = wdPreferredWidthPercent
    tbl.Columns(2).PreferredWidth = 85

    With tbl.Range.Font
        .Name = "Calibri"
        .Size = 10
    End With

    ' Remove borders
    With tbl.Borders
        .InsideLineStyle = wdLineStyleNone
        .OutsideLineStyle = wdLineStyleNone
    End With

    ' Fill table
    tbl.Cell(1, 1).Range.Text = "Account No:"
    tbl.Cell(1, 2).Range.Text = Nz(rsHeader!AccountNumber, "")

    tbl.Cell(2, 1).Range.Text = "Bank:"
    tbl.Cell(2, 2).Range.Text = Nz(rsHeader!BankName, "")

    tbl.Cell(3, 1).Range.Text = "Bank Address:"
    tbl.Cell(3, 2).Range.Text = Nz(rsHeader!BankAddress, "")

    tbl.Cell(4, 1).Range.Text = "Account Name:"
    tbl.Cell(4, 2).Range.Text = Nz(rsHeader!AccountName, "")

    tbl.Cell(5, 1).Range.Text = "ABA Routing Num:"
    tbl.Cell(5, 2).Range.Text = Nz(rsHeader!ABANumber, "")

    reText = "Monthly Lease Payments - Master Lease: " & _
        Nz(rsHeader!MLNo, "") & _
        " Schedule No: " & Nz(rsHeader!Schedule, "") & _
        " (" & Nz(rsHeader!GroupsIn, "") & ")"

    tbl.Cell(6, 1).Range.Text = "RE:"
    tbl.Cell(6, 2).Range.Text = reText

    ' Move cursor after table
    Set rng = WordDoc.Range(tbl.Range.End, tbl.Range.End)
    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd

    ' Grand Total
    rng.Text = "Grand Total to be transferred: " & Format(grandTotal, "$#,##0.##")
    rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
    rng.Font.Bold = True
    rng.Font.Size = 11

    rng.Collapse wdCollapseEnd
    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd

    '-------------------------
    ' TABLE 2 — RENT SCHEDULE
    '-------------------------
    rsDetail.MoveLast
    rsDetail.MoveFirst

    Set tbl = WordDoc.tables.Add(rng, rsDetail.RecordCount + 1, 8)
    tbl.Style = "Grid Table 1 Light"
    tbl.rows.Alignment = wdAlignRowLeft

    With tbl.Range.Font
        .Name = "Calibri"
        .Size = 10
    End With

    ' Blue header row
    With tbl.rows(1).Shading
        .BackgroundPatternColor = wdColorDarkBlue
    End With

    With tbl.rows(1).Range.Font
        .Color = wdColorWhite
        .Bold = True
    End With
    ' Alignment rules for table 2
    tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    tbl.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    tbl.Columns(1).Width = 90
    tbl.Columns(2).Width = 60
    tbl.Columns(3).Width = 90
    tbl.Columns(4).Width = 60
    tbl.Columns(5).Width = 60
    tbl.Columns(6).Width = 65
    tbl.Columns(7).Width = 100
    tbl.Columns(8).Width = 52

    ' Header row text
    rowIndex = 1
    With tbl
        .Cell(rowIndex, 1).Range.Text = "Assignee Contract No."
        .Cell(rowIndex, 2).Range.Text = "Periodic Basic Rent Payment Amount"
        .Cell(rowIndex, 3).Range.Text = "Lessee"
        .Cell(rowIndex, 4).Range.Text = "FA Sch No. (Group No.)"
        .Cell(rowIndex, 5).Range.Text = "*Taxes"
        .Cell(rowIndex, 6).Range.Text = "Amount " & vbNewLine & "to be Transferred (Rent+Taxes)"
        .Cell(rowIndex, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(rowIndex, 7).Range.Text = "**Rent Payment " & vbNewLine & " Due Dates"
        .Cell(rowIndex, 7).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(rowIndex, 8).Range.Text = "Payments Retained by Lessor"
    End With

    ' Data rows
    Do Until rsDetail.EOF
        rowIndex = rowIndex + 1

        faSchText = "Sch:" & Nz(rsDetail!Schedule, "") & vbCrLf & _
            "(Grp " & Nz(rsDetail!groupNo, "") & ")" & vbCrLf & _
            Nz(rsDetail!UnitsInGroup, 0) & _
            IIf(Nz(rsDetail!UnitsInGroup, 0) = 1, " unit", " units")

        With tbl
            .Cell(rowIndex, 1).Range.Text = Nz(rsDetail!Contract, "")
            .Cell(rowIndex, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cell(rowIndex, 2).Range.Text = Format(Nz(rsDetail!GroupRent, 0), "$#,##0.#0")
            .Cell(rowIndex, 3).Range.Text = Nz(rsDetail!clientCompanyName, "")
            .Cell(rowIndex, 4).Range.Text = faSchText
            .Cell(rowIndex, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(rowIndex, 5).Range.Text = "$"
            .Cell(rowIndex, 6).Range.Text = "$"
            .Cell(rowIndex, 7).Range.Text = Nz(rsDetail!RentDueDates, "")
            .Cell(rowIndex, 7).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Cell(rowIndex, 8).Range.Text = "N/A"
            .Cell(rowIndex, 8).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With

        rsDetail.MoveNext
    Loop

    '-------------------------
    ' FOOTER NOTES
    '-------------------------
    Set rng = WordDoc.Range(tbl.Range.End, tbl.Range.End)
    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd

    rng.Text = "*Monthly Taxes amount is subject to change should tax regulations change." & vbNewLine & _
        "** Basic Rent Due Dates are expected to be due [the 25th of each month] unless amended."
    rng.Font.TextColor = wdColorDarkBlue
    rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
    rng.Font.Size = 8

    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd

    rng.Text = "Acknowledged by:  REGB:______ " & _
        "BALCAP:______ " & _
        "FA:______"
    rng.Font.Size = 9
    rng.ParagraphFormat.Alignment = wdAlignParagraphRight
    rng.Font.TextColor = wdBlack
    rng.Font.Bold = True
    '-------------------------
    ' CLEANUP
    '-------------------------
    rsHeader.close
    rsDetail.close
    Set rsHeader = Nothing
    Set rsDetail = Nothing
    Set db = Nothing

    WordApp.Visible = True
    WordDoc.Activate

End Sub

Public Sub MakePAIOLD(iSchID As Long)

    DoCmd.close acReport, "PAIReport"

    SQLz = "Select "
    SQLz = SQLz & "ClientCompanyName"
    SQLz = SQLz & ",MLNo"
    SQLz = SQLz & ",Schedule"
    SQLz = SQLz & ",GroupNo"
    SQLz = SQLz & ",UnitsInGroup"
    SQLz = SQLz & ",MLDate"
    SQLz = SQLz & ",Lessee"
    SQLz = SQLz & ",LesseeAddress"
    SQLz = SQLz & ",GroupRent"
    SQLz = SQLz & ",RentDueDates"
    SQLz = SQLz & ",PayDate"
    SQLz = SQLz & ",BankShortName "
    SQLz = SQLz & ",ABANumber"
    SQLz = SQLz & ",AccountName"
    SQLz = SQLz & ",AccountNumber"
    SQLz = SQLz & ",BankName"
    SQLz = SQLz & ",BankAddress"
    SQLz = SQLz & ",BankRemitAddress"
    SQLz = SQLz & ",[Contract]"
    SQLz = SQLz & ",GroupsIn"
    SQLz = SQLz & ",ScheduleDate"
    SQLz = SQLz & ",SchID"
    'Sqlz = Sqlz & ",clientID"
    'Sqlz = Sqlz & ",SgrpId"
    'Sqlz = Sqlz & ",BankID"
    SQLz = SQLz & " INTO PAITemp1 From PAIBase WHERE SchID=" & iSchID & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.OutputTo acOutputReport, "PAIReport", acFormatPDF, "C:\FAAS\" & FAUserName & "\temp\PAIReport.pdf"
   'DoCmd.OutputTo acOutputReport, "PAIReport", acFormatRTF, "C:\FAAS\" & FAUserName & "\temp\PAIReport.rtf"
    DoCmd.SetWarnings True
    Set WordApp = CreateObject("Word.Application")
    Set WordDoc = WordApp.Documents.Add("C:\FAAS\" & FAUserName & "\temp\PAIReport.pdf")

    WordApp.Visible = True
    WordDoc.Activate
    For Each WordTable In WordDoc.tables
        WordTable.Select
        WordTable.AutoFormat
        WordTable.AutoFormat Format:=wdTableFormatClassic2
    Next
End Sub
Public Sub MakeRAL()

    Dim TempStr As String

    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From RALTemp1")
    DoCmd.OpenQuery "RALInsert"
    DoCmd.SetWarnings True

    Set RST = CurrentDb.OpenRecordset("SELECT * FROM RALTemp1")
    If RST!BankRemitAddress Like "*Banc of America*" Then
        SQLz = ""
        SQLz = SQLz & "Update RALTemp1"
        SQLz = SQLz & " Set"
        SQLz = SQLz & " Paragraph1 = Replace(Paragraph1, 'Assignee', 'Bank')"
        SQLz = SQLz & ",Paragraph2 = Replace(Paragraph2, 'Assignee', 'Bank')"
        SQLz = SQLz & ",Paragraph3 = Replace(Paragraph3, 'Assignee', 'Bank')"
        SQLz = SQLz & ",Paragraph4 = Replace(Paragraph4, 'Assignee', 'Bank')"
        SQLz = SQLz & ",Paragraph5 = Replace(Paragraph5, 'Assignee', 'Bank')"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    If DCount("SGrpID", "SchGrp", "sgrpSchID=" & RST!SchID & "") = 1 Then
        SQLz = ""
        SQLz = SQLz & "Update RALTemp1"
        SQLz = SQLz & " Set"
        SQLz = SQLz & " Paragraph4 = '<br/><br/><br/><br/><br/>' & Paragraph4"
        SQLz = SQLz & ",Paragraph5 = '<br />This letter may be executed in counterparts, which together shall constitute one and the same instrument.'"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
    End If
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputReport, "RALReport", acFormatPDF, "C:\FAAS\" & FAUserName & "\temp\RALReport.pdf"
    DoCmd.SetWarnings True
    Set WordApp = CreateObject("Word.Application")
    Set WordDoc = WordApp.Documents.Add("C:\FAAS\" & FAUserName & "\temp\RALReport.pdf")
    WordApp.Visible = True
    RST.close
    Set RST = Nothing
End Sub
Public Sub MakeRALAttachment()

    Dim SQLz, TempStr As String
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputReport, "RALAttachment", acFormatPDF, "C:\FAAS\" & FAUserName & "\temp\RALAttachment.pdf"
    DoCmd.SetWarnings True
    Set WordApp = CreateObject("Word.Application")
    Set WordDoc = WordApp.Documents.Add("C:\FAAS\" & FAUserName & "\temp\RALAttachment.pdf")
    WordApp.Visible = True
End Sub
Public Sub MakeRALAttachmentEXCEL(iSchID As Long)
    Dim RST As DAO.Recordset
    Dim GCount, i As Integer
    Dim Headerx, XLS As String

    SQLz = "Select CompanyName, MLNo, Schedule, GroupNo, SGrpID, SchID, GroupDesc "
    SQLz = SQLz & "FROM vw_RALAttachment "
    SQLz = SQLz & "WHERE SchID=" & iSchID & " Group By CompanyName, MLNo, Schedule, GroupNo, SGrpID, SchID, GroupDesc Order By GroupNO DESC"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    Call MessageUser("Open", "Making RAL EXCEL Attachement", "Working")

    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    Headerx = RST!CompanyName & " Schedule: " & RST!Schedule & " (" & RST!groupNo & ") Master Lease: " & RST!MLNo & ""
    GCount = RST.RecordCount
    Dim IsLastRpt As Boolean
    For i = 1 To GCount
        XLS = "|A,DD|A,cLabel:End=A:3:Unit" & vbNewLine & "Number|A,cColor:3:47.117.181|"
        XLS = XLS & "A,TX|"
        XLS = XLS & "B,cLabel:End=J:1:" & Headerx & "|"
        XLS = XLS & "C,cLabel:End=C:3:Year/Make/Model/Type" & vbNewLine & RST!GroupDesc & "|C,cColor:3:47.117.181|"
        XLS = XLS & "D,$E|D,cLabel:End=D:3:Lessor Cost" & vbNewLine & "(Per Unit)|D,cColor:3:47.117.181|"
        XLS = XLS & "E,$E|E,cLabel:End=E:3:Rent" & vbNewLine & "(Per Unit)|E,cColor:3:47.117.181|"
        XLS = XLS & "F,cColor:3:47.117.181|"
        XLS = XLS & "G,cColor:3:47.117.181|"
        XLS = XLS & "H,cColor:3:47.117.181|"
        XLS = XLS & "I,cColor:3:47.117.181|"
        XLS = XLS & "J,TC|J,cLabel:End=J:3:Delivery" & vbNewLine & "(*=Estimate)|J,cColor:3:47.117.181|"
        XLS = XLS & "K,TC|K,cLabel:End=K:3:Funded" & vbNewLine & "(Vendor Paid)|K,cColor:3:47.117.181|"
        XLS = XLS & "L,TR|L,cColor:3:47.117.181|"
        XLS = XLS & "M,cColor:3:47.117.181|"
        XLS = XLS & "A,LL|B,VAutoFit|"

        SQLz = "Select"
        SQLz = SQLz & " Unitnum"  '-Col A
        SQLz = SQLz & ",VIN"
        SQLz = SQLz & ",UnitDescDocVer" '-Col C
        SQLz = SQLz & ",LessorCost"
        SQLz = SQLz & ",UnitRent" '-Col E
        SQLz = SQLz & ",LAddress as Domicile"
        SQLz = SQLz & ",LCity as City"
        SQLz = SQLz & ",LState as State"
        SQLz = SQLz & ",LZip as Zip"
        SQLz = SQLz & ",DelvDate"
        SQLz = SQLz & ",FundedDate"
        SQLz = SQLz & ",BLCD"
        SQLz = SQLz & " FROM vw_RALAttachment "
        SQLz = SQLz & "WHERE SGrpID=" & RST!SGrpID & " Order By AssetID"
        If i = GCount Then IsLastRpt = True Else IsLastRpt = False

        Call SendReport2Excel(SQLz, RST!groupNo, Headerx, False, XLS, , , i, IsLastRpt, , "C:\Faas\" & FAUserName & "\Temp\RALAttachment.xlsx")
        RST.MoveNext
    Next i
    RST.close
    Set RST = Nothing
    Call MessageUser("Close")
End Sub
Public Sub makeword2()

    Dim WordHeaderFooter As HeaderFooter
    Dim RST As DAO.Recordset
    Set RST = CurrentDb.OpenRecordset("SELECT * FROM RALTemp1")
    Set WordApp = CreateObject("Word.Application")

    With WordApp
        .Visible = True

        Set WordDoc = .Documents.Add
        '   doc.SaveAs CurrentProject.path & "c:\FAAS\" & FAUserName & "\temp\TestDoc.doc"
    End With

    With WordApp.Selection

        .Font.Name = "Times New Roman"
        .Font.Size = 11

        .TypeText Format(Date, "mmmm dd, yyyy")
        .TypeParagraph
        .TypeText RST!LesseeAddress
        .TypeParagraph
        .TypeText RST!Paragraph1
        .TypeParagraph
        'Add header and footer

        '   ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = "Header"
        '   ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "Footer"
    End With
    RST.close
    'doc.Save
    'doc.Activate
End Sub
Public Function MakeCofALetter(cGroupID As Integer, Optional OtherReason As String)
    Dim Header, Paragraph1, Paragraph2, Paragraph3, Paragraph4, Paragraph5, Paragraph6, Paragraph7, cMLNo, SchedulesX, SQLx, cMLday, cMLMo, cMLYr As String
    Dim MstleaseLabel, SchLabel As String
    Dim Rst2    As DAO.Recordset
    Dim i, ii As Integer
    Dim cMLID, cSchID As Long

    cSchID = DLookup("SchId", "vw_SixKeys", "sgrpID=" & cGroupID)
    cMLID = DLookup("MLID", "vw_SixKeys", "sgrpID=" & cGroupID)
    SQLz = "SELECT * FROM MstrLease Where MLID=" & cMLID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    cMLNo = RST!MLNo
    cMLday = DatePart("D", RST!MLDt)
    cMLMo = Format(DatePart("m", RST!MLDt), "mmm")
    cMLYr = DatePart("YYYY", RST!MLDt)
    RST.close
    SQLz = ""

    Header = ""

    Paragraph1 = "1. The Equipment described in<u> Exhibit A</u> attached hereto and incorporated herein by this reference " & _
        "(hereinafter referred to as the “Exhibit A Equipment”) has been duly delivered to the location identified in <u>Exhibit A</u> and " & _
        "where it will be garaged or used."

    Paragraph2 = "2. All of the Exhibit A Equipment has been inspected and is determined to be (a) complete, (b) properly installed, (c) functioning, " & _
        "and (d) in good working order and in compliance with  all applicable  specifications."

    Paragraph3 = "3. The Exhibit A Equipment is of a size, design, capacity and manufacture acceptable to Lessee and suitable for Lessee's purposes."

    Paragraph4 = "4. Lessee acknowledges that signature of this document (a) constitutes unconditional acceptance of the Exhibit A Equipment under and " & _
        "subject to the terms of the Master Equipment Lease Agreement No. " & [cMLNo] & " dated as of " & [cMLMo] & " " & [cMLday] & ", " & [cMLYr] & " by and " & _
        "between Lessor and Lessee (as may be amended from time to time) and the Schedule thereunder relating to the Exhibit A Equipment, (b) that such acceptance is not on a trial basis, " & _
        "and (c) authorizes commencement of the Lease Term under the Master Equipment Lease Agreement and Schedule."

    Paragraph5 = "5. Lessee's execution of this Certificate of Acceptance with respect to the Exhibit A Equipment conclusively evidences that the Lessee has inspected and " & _
        "accepted each item of Exhibit A Equipment, and that the Lessee shall not reject, return or revoke this Agreement and any Schedule as to any such item of " & _
        "Exhibit A Equipment. Lessee will not contest the authority of any person executing a Certificate of Acceptance on Lessee's behalf. Each item of Exhibit A " & _
        "Equipment is leased in 'AS IS', 'WHERE IS' condition."

    Paragraph6 = "6. The Exhibit A Equipment shall at all times be the sole and exclusive property of Lessor and Lessee shall have no right, title or interest in or to the " & _
        "Exhibit A Equipment except the right to use the same upon the terms and conditions herein so long, and only so long, as Lessee shall not be in default in " & _
        "performance of any terms and conditions of the Master Equipment Lease Agreement, and Schedule or the Transaction Documents."

    Paragraph7 = "7. Lessee is not in default under the Master Equipment Lease Agreement or the Schedule, and all of Lessee's statements and promises contained in the Master " & _
        "Equipment Lease Agreement and Schedule remain true and correct."

    Header = Replace(Header, "'", "''")
    Paragraph1 = Replace(Paragraph1, "'", "''")
    Paragraph2 = Replace(Paragraph2, "'", "''")
    Paragraph3 = Replace(Paragraph3, "'", "''")
    Paragraph4 = Replace(Paragraph4, "'", "''")
    Paragraph5 = Replace(Paragraph5, "'", "''")
    Paragraph6 = Replace(Paragraph6, "'", "''")
    Paragraph7 = Replace(Paragraph7, "'", "''")

    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From CofATemp")
    SQLz = "Insert into CofATemp (Header, Paragraph1, Paragraph2, Paragraph3, Paragraph4, Paragraph5, Paragraph6, Paragraph7, MLID, SchID, GroupID) " & _
        "Select '" & Header & "','" & Paragraph1 & "','" & Paragraph2 & "','" & Paragraph3 & "','" & Paragraph4 & "','" & Paragraph5 & "','" & Paragraph6 & "','" & Paragraph7 & "'" & _
        "," & cMLID & "," & cSchID & "," & cGroupID & ";"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
End Function
Public Function MakeTitleRequestLetter(iBankId As Long, RequestNum As Integer, Optional OtherReason As String, Optional SaveFile As String, Optional NameToUse As String)
    Dim Header, Paragraph1, Paragraph2, Paragraph3, Paragraph4, Paragraph5, MLNox, SchedulesX, SQLx, BankX As String
    Dim MstleaseLabel, SchLabel As String
    Dim Rst2    As DAO.Recordset
    Dim i, ii As Integer
    Dim RAddr  As String
    Dim Committ As String

   'Get Custodion Address
   'Test for Wells Fargo Bank National Association / ABS Custody Vault Location
    Call MessageUser("Open", "Making the Requested Report", "This will take a moment...")
    SQLz = "Select * From Locations Where ClientRef=5233 AND LocationType='Billing';"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount > 0 Then
        RST.MoveFirst
        BankX = DLookup("BankName", "Banks", "BankId=" & iBankId)
        RAddr = RST!LocationStreet & "<br />"
        If Nz(RST!LAddrLine1, "") <> "" Then RAddr = RAddr & RST!LAddrLine1 & "<br />"
        If Nz(RST!LAddrLine2, "") <> "" Then RAddr = RAddr & RST!LAddrLine2 & "<br />"
        RAddr = RAddr & RST!LCity & ", " & RST!LState & " " & RST!LZip & "<br />"
        RAddr = Replace(RAddr, "SE MAC", "SE<br />MAC")
    Else
        RAddr = "No Custodion Address Set!"
        MsgBox RAddr, vbCritical, "No Custodion"
        RST.close
        Set RST = Nothing
        Call MessageUser("Close")
        Exit Function
    End If
    RST.close
    Set RST = Nothing
    'Set Header
    Header = "<p>" & Date & "<br /><br />"
    Header = Header & RAddr
    Header = Header & "<b>Attention</b>: Corporate Trust Services &ndash; Asset-Backed Securities Vault</p><br />"
    'Get Unit data
    SQLz = "Select MLNo "
    SQLz = SQLz & "From vw_SixKeys "
    SQLz = SQLz & "WHERE AssetID In (Select UnitID From Pick_Units_WorkTable) "
    SQLz = SQLz & "Group By MLNo "
    SQLz = SQLz & "Order By MLNo "
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveLast
    RST.MoveFirst
    Call MessageUser("Update", "Making the Requested Report", "Setting the data...")
    If RST.RecordCount = 1 Then MstleaseLabel = "Master Lease" Else MstleaseLabel = "Master Leases"
    MLNox = ""
    For i = 1 To RST.RecordCount
        SQLx = "Select Schedule From vw_SixKeys " & _
            "WHERE MLNo='" & RST!MLNo & "' and AssetID In (Select UnitID From Pick_Units_WorkTable) " & _
            "Group By Schedule Order By Schedule"

        Set Rst2 = CurrentDb.OpenRecordset(SQLx, dbOpenSnapshot)
        Rst2.MoveLast
        Rst2.MoveFirst
        If Rst2.RecordCount = 1 Then SchLabel = "Schedule" Else SchLabel = "Schedules"
        MLNox = MLNox & RST!MLNo & ": " & SchLabel & " ["
        SchedulesX = ""
        For ii = 1 To Rst2.RecordCount
            SchedulesX = SchedulesX & Rst2!Schedule & ", "
            Rst2.MoveNext
        Next ii
        SchedulesX = Left(SchedulesX, Len(SchedulesX) - 2)
        MLNox = MLNox & SchedulesX & "]<br />"
        RST.MoveNext
        Call MessageUser("Update", "Making the Requested Report", "Setting Sch data...")
    Next i
    '    SchedulesX = Left(SchedulesX, Len(SchedulesX) - 2) & "]"
    RST.close
    Rst2.close
    Set RST = Nothing
    Set Rst2 = Nothing
    If NameToUse & "test" = "test" Then NameToUse = FAUserFullName

    Paragraph1 = "<p>&#9;Reference is made to the Custody Agreement dated as of October 27, 2011 (as the same may be from time to time<br />" & _
        "amended, supplemented or otherwise modified; the <em>&ldquo;Custody Agreement&rdquo;</em>), by and between Fleet Advantage, LLC<br />(the<em> &ldquo;Company&rdquo;</em>)," & _
        "FA Title Trust (the <em>&ldquo;Trust&rdquo;</em>) and Wells Fargo Bank, National Association (the <em>&ldquo;Custodian&rdquo;</em>). <em>&nbsp;</em>Capitalized terms used herein shall " & _
        "have the respective meanings ascribed thereto in the Custody Agreement.</p>" & _
        "<br /><p>&#9;In connection with the administration of the Files held by you as the Custodian on behalf of the Owner " & _
        "Trustee, we request the release, and acknowledge receipt, of the (Equipment Files (Original Certificate(s) of Title) " & _
        "for the Equipment/Contract described below, for the reason indicated. " & _
        "<br /><br /><b>" & MstleaseLabel & "</b>: [<i>Schedule(s)</i>]:" & _
        "<br />" & MLNox & "" & _
        "<br /><b>Reason for Requesting Documents:</b><br />"

    If RequestNum = 1 Then _
        Paragraph2 = "<p>Lease of Equipment terminated (Servicer certifies the Lease of the Financed Vehicles has " & _
        "terminated).&nbsp; Release directly to:<b> " & NameToUse & " </b>at:<b> Fleet Advantage, LLC, " & _
        "401 East Las Olas Blvd., Suite 1720, Fort Lauderdale, FL 33301.</b></p>" & _
        "<p><br/><b>Note:</b><br />&#9;If all or part of the File was previously released to us, please release any additional documents in your" & _
        "<br/>&#9;possession relating to the specified Equipment.</p>"

    If RequestNum = 2 Then _
        Paragraph2 = "<p>Retitling (Servicer certifies that delivery of the original certificate of title is required in order to " & _
        "relocate and retitle the Financed Vehicle(s) in a new jurisdiction). " & _
        "Release directly to:<b> " & NameToUse & " </b>at:<b> Fleet Advantage, LLC, 401 East Las Olas Blvd., Suite 1720, " & _
        "Fort Lauderdale, FL 33301.</b></p>"
        ''& _
        ''"<p><br/><b>Note:</b><br />&#9;If all or part of the File was previously released to us, please release any additional documents " & _
        ''"<br/>&#9;in your possession relating to the specified Equipment.</p>"

    If RequestNum = 3 Then _
        Paragraph2 = "<p>The Financed Vehicle(s) has been purchased by the Beneficial Owner. Release directly to Release " & _
        "directly to:<b> " & NameToUse & " </b>at:<b> Fleet Advantage, LLC, 401 East Las Olas Blvd., Suite 1720, " & _
        "Fort Lauderdale, FL 33301.</b></p><br />"

    If RequestNum = 4 Then _
        Paragraph2 = "<p>Re-Registration (Servicer certifies that delivery of the original certificate of title is required in " & _
        "order to change or renew the registration of the Financed Vehicle(s). " & _
        "Release directly to: <b>" & NameToUse & " </b>at:<b> Fleet Advantage, LLC, 401 East Las Olas Blvd., <br />Suite 1720, " & _
        "Fort Lauderdale, FL 33301.</b></p><br />"


    If RequestNum = 7 Then _
        Paragraph2 = "<p><b>Other</b>: " & OtherReason & "<br />Release directly to:<b> " & NameToUse & "</b>at:<b> Fleet Advantage, LLC, 401 East Las Olas Blvd., Suite 1720, " & _
        "Fort Lauderdale, FL 33301.</b></p><br />"

    Paragraph3 = "<br /><p>&#9;This Request must be approved:&nbsp; by the Beneficial Owner if the Equipment constitutes General Trust " & _
        "Assets; by the applicable Holder if the Equipment constitutes Portfolio Assets; and by the applicable Lender if " & _
        "the Equipment is subject to a security interest in favor of such Lender and, only if the Custodian has received " & _
        "written notice from the Collateral Agent that an Event of Default under and as defined in the Backup Security Agreement " & _
        "has occurred, the Collateral Agent.</p><br />" & _
        "<p>&#9;The Custodian may conclusively rely on any Release Direction signed only by the Servicer and shall not " & _
        "be responsible for verifying the signature or identity of any Lender, Collateral Agent, Holder or Beneficial Owner, " & _
        "including any such signature required on a Transfer Request and shall be entitled to rely on the signatures " & _
        "provided on any Transfer Request as conclusive proof that the applicable consents have been obtained.</p>"

    Paragraph4 = "<p>&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;<u>FLEET ADVANTAGE, LLC</u><br />" & _
        "&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;Servicer <br /><br />" & _
        "&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;By: ________________________________<br />" & _
        "&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;Name:<b> " & NameToUse & "  </b><br />" & _
        "&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;&#9;Title:<b> Manager of Client Services Administration</b></p><br />"

    Paragraph5 = "APPROVED:<br /><br />" & _
        "FLEET ADVANTAGE, LLC<br /> " & _
        "Beneficial Owner<br /><br />" & _
        "By:________________________________<br />" & _
        "Name:<b> " & NameToUse & "  </b><br />" & _
        "Title:<b> Manager of Client Services Administration</b><br /><br /><br /><br />" & _
        "" & BankX & "<br />" & _
        "Holder<br /><br />" & _
        "By: _______________________________<br /><br />" & _
        "Name: ____________________________<br /><br />" & _
        "Title: _____________________________"

    Header = Replace(Header, "'", "''")
    Paragraph1 = Replace(Paragraph1, "'", "''")
    Paragraph2 = Replace(Paragraph2, "'", "''")
    Paragraph3 = Replace(Paragraph3, "'", "''")
    Paragraph4 = Replace(Paragraph4, "'", "''")
    Paragraph5 = Replace(Paragraph5, "'", "''")
    If iBankId = 4748 Then Paragraph5 = Replace(Paragraph5, "Bank & Trust", "Bank  & Trust")


    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From TitleRequestTemp")
    SQLz = "Insert into TitleRequestTemp (Header, Paragraph1, Paragraph2, Paragraph3, Paragraph4, Paragraph5) " & _
        "Select '" & Header & "','" & Paragraph1 & "','" & Paragraph2 & "','" & Paragraph3 & "','" & Paragraph4 & "','" & Paragraph5 & "';"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    If RequestNum = 5 Then
        If FAUserAccessLevel > 69 Then
            Committ = MsgBox("You're About to set the Lien Release RECV'D Date to: " & Date & " for the " & DCount("UnitID", "Pick_Units_WorkTable") & " Units you selected" & vbNewLine & "Do You Agree?", vbQuestion + vbYesNo + vbCritical, "Remove Units?")
            If Committ = vbYes Then
                SQLz = "UPDATE Units SET LienRelease='" & Date & "' WHERE UnitID in (Select UnitID From Pick_Units_WorkTable);"
                DoCmd.SetWarnings False
                DoCmd.RunSQL (SQLz)
                DoCmd.SetWarnings True
            End If
        End If
        DoCmd.OpenReport "TitleReleaseAffidavitOK", acViewPreview
    End If
    Call MessageUser("Update", "Making the Requested Report", "Final Data settings...")
    If RequestNum = 6 Then
        Committ = MsgBox("Do you want to Add the Title Event 'Sent Title Release Affidavit' for the " & DCount("UnitID", "Pick_Units_WorkTable") & " Unit(s) you selected" & vbNewLine & "Do You Agree?", vbQuestion + vbYesNo + vbCritical, "Add Event?")
        If Committ = vbYes Then
            Set RST = CurrentDb.OpenRecordset("Select UnitID FROM Pick_Units_WorkTable", dbOpenSnapshot)
            RST.MoveFirst
            For i = 1 To RST.RecordCount
                SQLz = "INSERT INTO UnitTitleLog ( UnitRef, EntryReason,  UnitTitleLogDate,  EntryTypeId,  EntryReasonId,  ExpectedReturn,  UserID )"
                SQLz = SQLz & "SELECT " & RST!UnitID & ", 'Sent Title Release Affidavit', #" & Date & "#, 2, 46, True, '" & FAUserName & "';"
                DoCmd.SetWarnings False
                DoCmd.RunSQL (SQLz)
                DoCmd.SetWarnings True
                RST.MoveNext
            Next i
        End If
        DoCmd.close acReport, "TitleReleaseAffidavit"
        'DoCmd.OpenReport "TitleReleaseAffidavit", acViewPreview
        DoCmd.SetWarnings False
        DoCmd.OutputTo acOutputReport, "TitleReleaseAffidavit", acFormatPDF, "C:\FAAS\" & FAUserName & "\temp\TitleReleaseAffidavit.pdf"
        DoCmd.SetWarnings True
        Set WordApp = CreateObject("Word.Application")
        Set WordDoc = WordApp.Documents.Add("C:\FAAS\" & FAUserName & "\temp\TitleReleaseAffidavit.pdf")
        WordApp.Visible = True
    End If

    If RequestNum < 5 Or RequestNum = 7 Then
        Dim folderPath As String
        Dim FilePath As String
        DoCmd.close acReport, "TitleReleaseRequest"
        If SaveFile <> "" Then
            folderPath = "C:\FAAS\" & FAUserName & "\temp\"
            FilePath = folderPath & SaveFile & ".pdf"
            Call MessageUser("Update", "Making the Requested Report", "Saving " & SaveFile & ".pdf")
            DoCmd.SetWarnings False
            DoCmd.OutputTo acOutputReport, "TitleReleaseRequest", acFormatPDF, FilePath
            DoCmd.SetWarnings True
            Call MessageUser("Update", "Making the Requested Report", "Opening " & SaveFile & ".pdf")
            Shell "cmd /c start msedge.exe """ & FilePath & """", vbNormalFocus
            'Shell "explorer.exe """ & folderPath & """, vbNormalFocus"
            'FollowHyperlink "C:\FAAS\" & FAUserName & "\temp\" & SaveFile & ".pdf"
        Else
            DoCmd.OpenReport "TitleReleaseRequest", acViewPreview
        End If
    End If
    Call MessageUser("Close")
End Function
