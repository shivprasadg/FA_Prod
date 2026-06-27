Attribute VB_Name = "GetExcelData"
Option Compare Database
Option Base 1
Private m_sNameOfOutPutWorkSheet_1 As String
Function OpenInspectionFileDialog(InspectionNum As Integer)
    Dim f      As Object
    Set f = Application.FileDialog(3)
    ' f.InitialFileName = "P:\FLEETADVANTAGE\CUSTOMERS"
    f.InitialFileName = "https://fleetadvantage.sharepoint.com/Customers/AllItems.aspx"
    f.AllowMultiSelect = False
    f.Filters.Clear
    f.Filters.Add "Excel", "*.xlsx"
    f.Filters.Add "Excel 97", "*.xlx"
    f.Filters.Add "CSV", "*.csv"
    f.Show
    On Error GoTo FixMe
    ' MsgBox "file choosen = " & f.SelectedItems(1)
    Call GetInspectionDataExcel(f.SelectedItems(1), InspectionNum)
    Exit Function
FixMe:
End Function
Public Function DownloadNISData(Optional URL As String, Optional FileNameX As String)
If URL = "" Then URL = "https://www.imanpro.net/admin/login.php?co=nis&ckey=CRPmDXodgNhKEAOVESXYL1q5Ipa1vNf7Xuou3sPIZensauM8pc7ve8veKSgSHFqm"
 Dim WebString As String
 Dim TableString As String
OpenEdge:

    Dim wb As WebDriver
    Set wb = New Selenium.EdgeDriver
    'Set WB = New Selenium.PhantomJSDriver
    wb.start
    wb.Get URL
Login:
    'Call MessageUser("Update", "Loging into NIS", "this will Go Quickly")
   wb.FindElementByXPath("/html/body/div[2]/form/div/div[2]/a/i").Click
   Pause (2)
   wb.FindElementByXPath("//*[@id=""limSelect""]/option[7]").Click
   WebString = wb.ExecuteScript("return document.body.innerHTML;")
   TableString = "<table>" & SplitString(SplitString(WebString, "<thead><tr>", 1), "</table>", 0) & "/Table>"

    Dim intFileNum As Integer
    intFileNum = FreeFile
    ' change Output to Append if you want to add to an existing file
    ' rather than creating a new file each time
    Open "C:\Faas\" & FAUserName & "\Temp\WebText.txt" For Output As intFileNum
    Print #intFileNum, TableString
    Close intFileNum
SetDate:
End Function

Public Function DownloadDekraDataExcel(Optional URL As String, Optional FileNameX As String)

    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim eHTMLElement As IHTMLElement
    Dim LastImportedInsp As Long
    Dim WebString As String
    Dim WebSection As String
    Dim WebSource As String
    Dim SectionX As String
    Dim SelectedRange, WorkingRange As Range
    Dim i, x   As Integer
    Dim TargetX, SQLz, DataX() As String
    'File Finder stuff
    Dim iFilesNum As Integer
    Dim iCount As Integer
    Dim sPathArry() As String
    Dim iFilesFound As Boolean
    Dim sPath  As String
    Dim sFile  As String
    Dim sFile2  As String
    Dim sSearch As String
    Dim sStartDate As String
    On Error GoTo 0
    If URL = "" Then URL = "https://www.dekrasmart.com/Analysis/Login"
    'sStartDate = DatePart("M", (Date - 15)) & "/1/" & DatePart("yyyy", Date)
    sStartDate = Date - 5
    sPath = "C:\Users\" & FAUserName & "\Downloads\"
    sSearch = "Schedule_" & DatePart("YYYY", Date) & "-"
    If DatePart("M", Date) < 10 Then sSearch = sSearch & "0" & DatePart("M", Date) & "-" Else sSearch = sSearch & DatePart("M", Date) & "-"
    If DatePart("D", Date) < 10 Then sSearch = sSearch & "0" & DatePart("D", Date) Else sSearch = sSearch & DatePart("D", Date)
    sSearch = sSearch & "_*.xlsx"
    If FileNameX = "" Then iFilesFound = FindFiles(sPath, sPathArry, iFilesNum, sSearch, True)

    Call MessageUser("Open", "Gettng Dekra Inspections", "This might take a bit...")

KillPreviousfile:
If FileNameX = "" Then
    If iFilesFound Then
        For iCount = 1 To iFilesNum
            Call MessageUser("Update", "Deleting Previous files:", sPathArry(2, iCount))
            Kill (sPathArry(1, iCount) & sPathArry(2, iCount))
        Next
    End If
    DoCmd.SetWarnings True
    Call MessageUser("Update", "Opening Dekra's Site", "Relax... it will take a moment.")
OpenEdge:

    Dim wb As WebDriver
    Set wb = New Selenium.EdgeDriver
    'Set WB = New Selenium.PhantomJSDriver
    wb.start
    wb.Get URL
Login:
    Call MessageUser("Update", "Loging into Dekra", "this will Go Quickly")
    wb.FindElementById("Login").SendKeys "ljohnson@fleetadvantage.com"
    wb.FindElementById("Senha").SendKeys "3kyv#1"
    wb.FindElementByClass("btn-success").Click
    wb.FindElementById("tipoPeriodo").SendKeys "Completed"
    wb.FindElementById("DataInicio").SendKeys sStartDate
    wb.FindElementById("btnConsultar").Click
SetDate:
    Call MessageUser("Update", "Setting date Search", "Getting this month's data")
    wb.Wait (1000)
SetandGetDownload:
    wb.FindElementById("btnExportarExcelDados").Click

    wb.Wait (1000)
    wb.FindElementById("summary").Click
    wb.Wait (1000)
    'disabled Locations 7/19/2022 - Leo per CV
    'WB.FindElementById("location").Click
    'WB.Wait (1000)
    wb.FindElementById("toMaxResolution").Click
    wb.Wait (1000)
    wb.FindElementByXPath("/html/body/div[14]/div[3]/div/button[1]").Click
    wb.Wait (4000)
CloseEdge:
    wb.close
    Set edge = Nothing
GetFile:
    sFile = ""
    iFilesFound = FindFiles(sPath, sPathArry, iFilesNum, sSearch, True)
    If iFilesFound Then
        For iCount = 1 To iFilesNum
           ' MsgBox "Downloaded File: " & sPathArry(2, iCount), vbInformation, "Success!"
            sFile = sPathArry(2, iCount)
        Next
    End If

SetPaths:
    sFile = sPath & sFile
    Else
    sFile = FileNameX
    End If

    Call MessageUser("Update", "Copying Excel to Archive", "Inspection Import will Start Soon")

    sFile2 = "\\FA-Poseidon\Atlaas\Inspections\Archive\Imported_" & DatePart("M", Date) & "_" & DatePart("D", Date) & "_" & DatePart("yyyy", Date) & ".xlsx"
    FileCopy sFile, "\\FA-Poseidon\Atlaas\Inspections\NewDekraInspections.xlsx"
    FileCopy sFile, sFile2
    sFile = "\\FA-Poseidon\Atlaas\Inspections\NewDekraInspections.xlsx"

    Call MessageUser("Close")

    Call CleanDekraImport(sFile, Date - 10, Date, True)
    Call GetDekraInspections
    Call SendDekraImportEmail(sFile2)

End Function


Sub findFile(sPath As String, sSearch As String)
    Dim i      As Integer
    Dim fs     As FileSearch
    Dim numFiles As Long
    Set fs = Application.FileSearch

    With fs
        .NewSearch
        .LookIn = sPath
        .fileName = sSearch
        .SearchSubFolders = True
        If .Execute() > 0 Then
            For i = 1 To .FoundFiles.Count
                Debug.Print .FoundFiles(i)
            Next i
        End If
    End With
End Sub
Public Function GetDekraInspections()

    Dim i, x, InspCount   As Integer
    Dim TargetX, SQLz, UTypex, DataX() As String
    Dim DekraDatax As String
    On Error GoTo 0
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From InspLinkTemp")
    DoCmd.SetWarnings True
    Call MessageUser("Open", "Opening Inspections", "This might take a bit...")
    DoCmd.SetWarnings True
    Dim RST    As DAO.Recordset
    Set RST = CurrentDb.OpenRecordset("Select * From NewDekraInspectionsImport", dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    If RST.RecordCount = 0 Then
        MsgBox "There aren't any new inspections" & vbNewLine & "Nothing was imported", vbInformation, "We're Up To Date"
        Call MessageUser("Close")
        Exit Function
    End If
    InspCount = RST.RecordCount
    For i = 1 To InspCount
        DekraDatax = GetDekraData(RST!LinkAddr)
        Call MessageUser("Update", "ODO Re:Insp " & i & " of:" & InspCount & ", UnitID: " & RST!UnitID & "")
        'UTypex = DLookup("UnitType", "Units", "UnitID=" & RST!UnitID)
        SQLz = ""
        SQLz = SQLz & "Insert into InspLinkTemp ("
        SQLz = SQLz & " UnitID "
        SQLz = SQLz & ",InspCo "
        SQLz = SQLz & ",Link2 "
        SQLz = SQLz & ",UnitNum "
        SQLz = SQLz & ",INSP "
        SQLz = SQLz & ",iCost "
        SQLz = SQLz & ",VIN "
        SQLz = SQLz & ",InspDate "
        SQLz = SQLz & ",Odometer "
        SQLz = SQLz & ",iHourRead "
        SQLz = SQLz & ") SELECT "
        SQLz = SQLz & RST!UnitID
        SQLz = SQLz & ", 132"
        SQLz = SQLz & " ,'Dekra#" & RST!LinkAddr & "#'"
        SQLz = SQLz & " ,'" & RST!iUnitNum & "'"
        SQLz = SQLz & " ," & RST!INSP & ""
        SQLz = SQLz & " ," & RST!iCost & ""
        SQLz = SQLz & " ,'" & RST!VIN & "'"
        SQLz = SQLz & ",#" & CDate(SplitString(DekraDatax, "|", 0)) & "#"
        SQLz = SQLz & ",  " & CLng(SplitString(DekraDatax, "|", 1)) & ""
        SQLz = SQLz & ",  " & CLng(SplitString(DekraDatax, "|", 2)) & ""
        SQLz = SQLz & ";"
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLz)
        DoCmd.SetWarnings True
        RST.MoveNext
    Next i
InsertInsp:
    Call MessageUser("Update", "Inserting...", "Inspections, If any are new")
    SQLz = "INSERT INTO UnitInspections (AssetID, INSP, iDate, iOdoRead, iResultID, iNotes, iVendorID, InspNum, iWebLink, ItypeID, iCost,iHourRead ) "
    SQLz = SQLz & "SELECT InspLinkTemp.UnitId, InspLinkTemp.INSP, InspLinkTemp.InspDate, InspLinkTemp.Odometer, 3 AS ResultID "
    If DekraDatax Like "*QUICK*" Then
    SQLz = SQLz & ",'Dekra Insp Data Imported on " & Date & "' AS Notes"
    Else
    SQLz = SQLz & ",'Dekra Insp Data Imported on " & Date & " QUICK Inspection ONLY' AS Notes"
    End If
    SQLz = SQLz & ",InspLinkTemp.InspCo, Nz([ix],0)+1 AS InspNum, InspLinkTemp.Link2, 1, iCost, iHourRead "
    SQLz = SQLz & "FROM InspLinkTemp "
    SQLz = SQLz & "LEFT JOIN "
    SQLz = SQLz & "(SELECT AssetID, Max(UnitInspections.InspNum) AS ix FROM UnitInspections GROUP BY AssetID) AS L "
    SQLz = SQLz & "ON InspLinkTemp.UnitID = L.AssetID;"

    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

AddExpense:
    Call MessageUser("Update", "Inserting RMKT Expense", "For the new inspections")
    SQLz = "INSERT INTO RemarketExpense ( AssetId, Expense, ExpenseDate, ExpenseTypeId, VendorId, Comment )"
    SQLz = SQLz & "SELECT InspLinkTemp.UnitId"
    SQLz = SQLz & ", InspLinkTemp.icost"
    SQLz = SQLz & ", InspLinkTemp.InspDate"
    SQLz = SQLz & ", 14 AS Type"
    SQLz = SQLz & ", 132 AS VendorID"
    SQLz = SQLz & ", " & Chr(34) & "INSP: " & Chr(34) & " & [INSP] & " & Chr(34) & ", Imported by Atlass on: " & Chr(34) & " & Date() AS Comment "
    SQLz = SQLz & "FROM InspLinkTemp LEFT JOIN "
    SQLz = SQLz & "(SELECT VendorId, ExpenseTypeId, Expense, ExpenseDate, AssetId "
    SQLz = SQLz & "FROM RemarketExpense  as E INNER JOIN InspLinkTemp as I  ON E.AssetId = I.UnitID "
    SQLz = SQLz & "WHERE E.VendorId=132 AND E.ExpenseTypeId=14)  AS ExpIn ON InspLinkTemp.UnitID = ExpIn.AssetId "
    SQLz = SQLz & "WHERE ExpIn.AssetId Is Null;"

    Call MessageUser("Update", "Inserting Expenses...", "If any are new")
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Call MessageUser("Close")
    RST.close
    Set RST = Nothing
End Function

Public Function GetDekraData(sLink As String) As String
    'If sLink Like "*#*" Then sLink = SplitString(SplitString(sLink, "#", 1), "#", 0)
    Dim Tempx, UTypex, Hoursx, ODOx, InspDatex As String
    Dim wb As WebDriver
    Set wb = New Selenium.PhantomJSDriver
    wb.start
    wb.Get sLink
    InspDatex = wb.FindElementByXPath("/html/body/div[1]/div[1]/fieldset[1]/div[2]/div[4]/span").Text
    Hoursx = "0"
    ODOx = "0"
    On Error Resume Next
    Tempx = wb.FindElementByXPath("/html/body/div[1]/div/fieldset[1]/div[2]/div[1]").Text
    On Error GoTo 0
    If Not Tempx Like "*QUICK CHECK INSPECTION*" Then
        If Tempx = "" Then Tempx = wb.FindElementByXPath("/html/body/div[1]/div[2]/table/tbody/tr[2]/td[2]").Text
        If Tempx Like "*Trailer*" Or Tempx Like "*CX-5307*" Then UTypex = "Trailer"
        If Tempx Like "*TRUCK WITH BODY*" Then UTypex = "Straight" Else UTypex = "Powered"
        If Tempx Like "*Reefer*" Or Tempx Like "*Refer*" Then UTypex = UTypex & "|Reefer"
        If UTypex = "Powered|Reefer" Then UTypex = "Straight|Reefer"
        If UTypex Like "*Trailer*" Then
            ODOx = "0"
            If UTypex Like "*Reefer*" Then Hoursx = wb.FindElementByXPath("/html/body/div[1]/div[2]/fieldset/div[6]/div[7]/div[2]/label").Text
        Else
            If Hoursx = "" And UTypex = "Trailer|Reefer" Then Hoursx = wb.FindElementByXPath("/html/body/div[1]/div[2]/fieldset/div[6]/div[6]/div[2]/label").Text
            If Tempx Like "*TRUCK WITH BODY AND*" Then
                ODOx = SplitString(wb.FindElementByXPath("/html/body/div[1]/div[2]/div[1]/table/tbody/tr[9]/td[2]").Text, " ", 0)
                Hoursx = SplitString(wb.FindElementByXPath("/html/body/div[2]/div[2]/fieldset/div[22]/div/div[2]/label").Text, " ", 0)
            Else
               If Not Tempx Like "*Trailer*" Then
                   ODOx = SplitString(wb.FindElementByXPath("/html/body/div[1]/div[2]/table/tbody/tr[9]/td[2]").Text, " ", 0)
                   Hoursx = SplitString(SplitString(wb.FindElementByXPath("/html/body/div[1]/div[3]/table/tbody/tr[28]/td[2]").Text, "(", 1), ")", 0)
               End If
            End If
        End If
    Else
        UTypex = Tempx
        ODOx = "0"
        Hoursx = "0"
    End If
GotIT:
    If Not IsNumeric(ODOx) Then ODOx = "0"
    If Not IsNumeric(Hoursx) Then Hoursx = "0"
    GetDekraData = InspDatex & "|" & ODOx & "|" & Hoursx & "|" & UTypex
    wb.Quit
    Set wb = Nothing
End Function

Public Function CleanDekraImport(sFile As String, dStartDate As String, dEndDate As String, ImportFile As Boolean)
    DoCmd.SetWarnings False
    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim xlWB   As Object
    Dim SelectedRange, WorkingRange As Range
    Dim i, x, UnitYearX, xInspCnt As Integer
    Dim SQLz, DataX(), VinX, XLS, Headerx, LinkAddr As String
    Dim lastRow As Long
    Dim TotalRows As Long
    Dim RST    As DAO.Recordset
    Call MessageUser("Open", "Cleaning Inspections", "...this should be quick...")
    SQLz = "Select Insp, iDate From UnitInspections Where Insp Is Not Null and iDate >= #" & dStartDate & "# And [iDate] <= #" & dEndDate & "#;"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    RST.MoveLast
    RST.MoveFirst
    On Error GoTo 0
    xInspCnt = RST.RecordCount
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = True
        Set xlWB = .Workbooks.Open(sFile, , False)
        With .ActiveSheet    ' Remove dups, get data rows from source
StartLooking:
            Call MessageUser("Update", "Checking For Dups", "...the original file will be changed...")
            lastRow = .Range("A" & .rows.Count).End(xlUp).Row
            .Application.Range("T1") = "LinkAddr"
GetLinkAddr:
            For i = 2 To lastRow
                .Application.Range("P" & i).Select
                LinkAddr = Nz(.Application.Selection.Hyperlinks.Item(1).Address, "")
                If LinkAddr <> "" Then .Application.Range("T" & i) = LinkAddr
            Next i
            GoTo Finalize
RemoveDups:
            If xInspCnt > 0 Then
                For i = 2 To lastRow
                    For x = 1 To xInspCnt

                        .Application.Range("Q" & i).Select      'Check If Inspec in already imported
                        If .Application.Selection.Text <> "" Then
                            If CLng(.Application.Selection.Text) = RST!INSP Then
                                .Application.Range("Q" & i).EntireRow.Delete
                                RST.MoveFirst
                                Call MessageUser("Update", "Found Dup", "Insp:" & RST!INSP & " was already Imported")
                                GoTo StartLooking
                            End If
                        End If
                        RST.MoveNext
                    Next x
                    RST.MoveFirst
                Next i
            End If
            'xlApp.Application.Range("A1") = "VIN"
            'xlApp.Application.Range("B1") = "UnitYear"
Finalize:
            xlApp.Application.DisplayAlerts = False
            xlApp.ActiveWorkbook.SaveAs fileName:=sFile
            xlApp.Application.ActiveWorkbook.Save
            xlApp.Application.DisplayAlerts = True
            xlApp.Application.ActiveWorkbook.close
            xlApp.Application.Quit
            Call MessageUser("Update", "Copying Excel to Archive", "Inspection Import will Start Soon")
         '   FileCopy sFile, "A:\Inspections\NewDekraInspections.xlsx"
         '   FileCopy sFile, "A:\Inspections\Archive\Imported_" & DatePart("M", Date) & "_" & DatePart("D", Date) & "_" & DatePart("yyyy", Date) & ".xlsx"
            Call MessageUser("Close")
        End With
    End With
    'If ImportFile = True Then Call GetInspectionDataExcel("A:\Inspections\NewDekraInspections.xlsx", 1)
End Function
Public Function GetInspectionDataExcel(sFile As String, ByVal InspectionNum As Integer)

    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim ie     As Object
    Dim WebString As String
    Dim WebSection As String
    Dim WebSource As String
    Dim SectionX As String
    Dim SelectedRange, WorkingRange As Range
    Dim i, x   As Integer
    Dim TargetX, SQLz, DataX() As String
    On Error GoTo 0
    Call MessageUser("Open", "Opening Inspections", "This might take a bit...")
    DoCmd.SetWarnings True
    'GoTo UpdateX

    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete From InspLinkTemp")
    DoCmd.SetWarnings True
    Set xlApp = CreateObject("Excel.Application")
    'mFile = sFile
    If sFile Like "*.xlsx" Then sFile = sFile Else sFile = sFile & ".xlsx"
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = False
        Set xlWB = .Workbooks.Open(sFile, , False)
        Dim lastRow As Long
        With .ActiveSheet
            lastRow = .Range("D" & .rows.Count).End(xlUp).Row
            If StartRow > 1 Then lastRow = lastRow + (StartRow - 1)
        End With
        '   ReDim DataX(1 To LastRow, 1 To 5)
        Set ie = CreateObject("InternetExplorer.Application")
        edge.Visible = False
        For i = 2 To lastRow
            .Application.Range("B" & i).Select
            If .Application.Selection.Text = "completed" Then
                .Application.Range("P" & i).Select
                edge.Navigate .Application.Selection.Hyperlinks.Item(1).Address
                While edge.Busy
                    DoEvents
                Wend
                WebString = edge.Document.DocumentElement.outerHTML
                If WebString Like "*NTP Engine*" Then WebSource = "NTP" Else WebSource = "Dekra"
                SQLz = ""
                SQLz = SQLz & "Insert into InspLinkTemp ("
                SQLz = SQLz & "InspCo "
                SQLz = SQLz & ",Link2 "
                SQLz = SQLz & ",UnitNum "
                SQLz = SQLz & ",INSP "
                SQLz = SQLz & ",iCost "
                SQLz = SQLz & ",VIN "
                SQLz = SQLz & ",InspDate "
                SQLz = SQLz & ",Odometer "
                SQLz = SQLz & ") SELECT "
                SQLz = SQLz & "132 "
                SQLz = SQLz & ",'Dekra#" & .Application.Selection.Hyperlinks.Item(1).Address & "#'"

                SectionX = "UnitNum"
                .Application.Range("C" & i).Select
                SQLz = SQLz & ",'" & .Application.Selection.Text & "'"

                SectionX = "INSP"
                .Application.Range("Q" & i).Select
                SQLz = SQLz & ",'" & .Application.Selection.Text & "'"

                SectionX = "iCost"
                .Application.Range("R" & i).Select
                SQLz = SQLz & "," & .Application.Selection.Text & ""

                SectionX = "VIN"
                .Application.Range("K" & i).Select
                SQLz = SQLz & ",'" & .Application.Selection.Text & "'"

                If WebSource = "Dekra" Or WebSource = "NTP" Then
                    SectionX = "InspDate"
                    WebSection = SplitString(WebString, "FLEET NAME</label>", 0)
                    WebSection = Trim(SplitString(WebSection, "<span>", 4))
                    WebSection = Trim(SplitString(WebSection, "</span>", 0))
                    WebSection = Replace(WebSection, vbCrLf, "")
                    WebSection = Replace(WebSection, vbNewLine, "")
                    WebSection = Replace(WebSection, Chr(10), "")
                    If Not IsDate(WebSection) Then
                        WebSection = Trim(SplitString(WebString, ">VIN #:", 0))
                        WebSection = Trim(SplitString(WebSection, "<span class=" & Chr(34) & "resposta" & Chr(34) & ">", 2))
                        WebSection = Trim(SplitString(WebSection, "</span>", 0))
                        WebSection = Replace(WebSection, vbCrLf, "")
                        WebSection = Replace(WebSection, vbNewLine, "")
                        WebSection = Replace(WebSection, Chr(10), "")
                    End If
                    If Not IsDate(WebSection) Then WebSection = "1/1/1900"
                    SQLz = SQLz & ",#" & Trim(WebSection) & "#"

                    SectionX = "InspODO"
                    WebSection = ""
                    WebSection = Trim(SplitString(WebString, " (ODOM)", 0))
                    WebSection = Trim(SplitString(WebSection, "Mileage:", 1))
                    WebSection = Trim(SplitString(WebSection, ">", 3))
                    If WebSection Like "*KM*" Then
                        WebSection = GetNumbers(WebSection)
                        If IsNumeric(WebSection) Then WebSection = Round(CLng(WebSection) * 0.62137119)
                    End If
                    If Not IsNumeric(WebSection) Then
                        WebSection = Trim(SplitString(WebString, ">ECU Mileage:</label>", 0))
                        WebSection = Trim(SplitString(WebSection, ">Odometer:</label>", 1))
                        WebSection = GetNumbers(WebSection)

                    End If
                    If Not IsNumeric(WebSection) Then
                        WebSection = Trim(SplitString(WebString, ">ECM Total Miles:", 0))
                        WebSection = Trim(SplitString(WebSection, ">Odometer:</span>", 1))
                        WebSection = Trim(SplitString(WebSection, "resposta", 1))
                        WebSection = Trim(SplitString(WebSection, "</span>", 0))
                        WebSection = GetNumbers(WebSection)
                    End If
                    If Not IsNumeric(WebSection) Then WebSection = 0
                    ' If WebSection Like "*Error*" Then WebSection = 0
                    SQLz = SQLz & "," & Replace(WebSection, ",", "") & ""
                Else
                    SectionX = "InspDate"
                    WebSection = ""
                    WebSection = Trim(SplitString(WebString, ">VIN #:", 0))
                    WebSection = Trim(SplitString(WebSection, "<span class=" & Chr(34) & "resposta" & Chr(34) & ">", 2))
                    WebSection = Trim(SplitString(WebSection, "</span>", 0))
                    WebSection = Replace(WebSection, vbCrLf, "")
                    WebSection = Replace(WebSection, vbNewLine, "")
                    WebSection = Replace(WebSection, Chr(10), "")
                    If WebSection Like "*Error*" Or Nz(WebSection, "") = "" Then WebSection = "1/1/1900"
                    SQLz = SQLz & ",#" & Trim(WebSection) & "#"
                End If

                SQLz = SQLz & ";"
                SectionX = "InsertTemp"
                Call MessageUser("Update", "Working On Row " & i - 1 & " of " & lastRow - 1 & "", .Application.Selection.Text)
                DoCmd.SetWarnings False
                DoCmd.RunSQL (SQLz)
                DoCmd.SetWarnings True

            End If
        Next i
        edge.Quit
        Set ie = Nothing
    End With
    xlApp.Application.Quit
    SQLz = ""
UpdateX:
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("UPDATE InspLinkTemp INNER JOIN Units ON InspLinkTemp.VIN = Units.UnitVIN and InspLinkTemp.UnitNum = Units.UnitUnitNum  SET InspLinkTemp.UnitID = [Units].[UnitID];")
    DoCmd.SetWarnings True

    Call MessageUser("Update", "Updating...", "Updating Existing Inspections, If any")

    SQLz = "UPDATE InspLinkTemp "
    SQLz = SQLz & "INNER JOIN UnitInspections ON InspLinkTemp.UnitID = UnitInspections.AssetID "
    SQLz = SQLz & "SET UnitInspections.iOdoRead = [InspLinkTemp].[Odometer] "
    SQLz = SQLz & ", UnitInspections.iWebLink = [InspLinkTemp].[Link2] "
    SQLz = SQLz & ", UnitInspections.iDate = [InspLinkTemp].[InspDate] "
    SQLz = SQLz & ", UnitInspections.iVendorID = [InspLinkTemp].[InspCo] "
    SQLz = SQLz & ", UnitInspections.iResultID = 3 "
    SQLz = SQLz & "WHERE UnitInspections.iResultID=4"


    SectionX = "UpdateRMKT"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Call MessageUser("Update", "Inserting...", "Inspections, If any are new")

    SQLz = "INSERT INTO UnitInspections (AssetID, INSP, iDate, iOdoRead, iResultID, iNotes, iVendorID, InspNum, iWebLink, ItypeID, iCost ) "
    SQLz = SQLz & "SELECT InspLinkTemp.UnitId, InspLinkTemp.INSP, InspLinkTemp.InspDate, InspLinkTemp.Odometer, 3 AS ResultID "
    SQLz = SQLz & ",'Dekra Insp Data Imported on " & Date & "' AS Notes"
    SQLz = SQLz & ",InspLinkTemp.InspCo, Nz([ix],0)+1 AS InspNum, InspLinkTemp.Link2, 1, iCost "
    SQLz = SQLz & "FROM InspLinkTemp "
    SQLz = SQLz & "LEFT JOIN "
    SQLz = SQLz & "(SELECT AssetID, Max(UnitInspections.InspNum) AS ix FROM UnitInspections GROUP BY AssetID) AS L "
    SQLz = SQLz & "ON InspLinkTemp.UnitID = L.AssetID;"

    SectionX = "InsertRMKT"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
AddExpense:
    SQLz = "INSERT INTO RemarketExpense ( AssetId, Expense, ExpenseDate, ExpenseTypeId, VendorId, Comment )"
    SQLz = SQLz & "SELECT InspLinkTemp.UnitId"
    SQLz = SQLz & ", InspLinkTemp.icost"
    SQLz = SQLz & ", InspLinkTemp.InspDate"
    SQLz = SQLz & ", 14 AS Type"
    SQLz = SQLz & ", 132 AS VendorID"
    SQLz = SQLz & ", " & Chr(34) & "INSP: " & Chr(34) & " & [INSP] & " & Chr(34) & ", Imported by Atlass on: " & Chr(34) & " & Date() AS Comment "
    SQLz = SQLz & "FROM InspLinkTemp LEFT JOIN "
    SQLz = SQLz & "(SELECT VendorId, ExpenseTypeId, Expense, ExpenseDate, AssetId "
    SQLz = SQLz & "FROM RemarketExpense  as E INNER JOIN InspLinkTemp as I  ON E.AssetId = I.UnitID "
    SQLz = SQLz & "WHERE E.VendorId=132 AND E.ExpenseTypeId=14)  AS ExpIn ON InspLinkTemp.UnitID = ExpIn.AssetId "
    SQLz = SQLz & "WHERE ExpIn.AssetId Is Null;"

    Call MessageUser("Update", "Inserting Expenses...", "If any are new")
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call MessageUser("Close")
    Exit Function
FixMe:
    edge.Quit
    Set ie = Nothing
    xlApp.Application.Quit
    Call MessageUser("Close")
    MsgBox "There was an error..." & vbNewLine & "Erorr Code: " & Err.Number & " (" & Err.Description & ")" & vbNewLine & " Code Section: " & SectionX & " Item:" & i, vbCritical, "ATLAAS Info"
End Function

Public Function GetInspectionData(URL As String)
    Dim ie     As Object
    Dim WebString As String
    Set ie = CreateObject("InternetExplorer.Application")
    edge.Visible = False
    edge.Navigate URL
    While edge.Busy
        DoEvents
    Wend
    WebString = edge.Document.DocumentElement.outerHTML
    edge.Quit
    Set ie = Nothing
    Dim intFileNum As Integer
    intFileNum = FreeFile
    ' change Output to Append if you want to add to an existing file
    ' rather than creating a new file each time
    Open "C:\Faas\" & FAUserName & "\Temp\WebText.txt" For Output As intFileNum
    Print #intFileNum, WebString
    Close intFileNum
End Function
Public Function ExtractDate(Value As Variant) As Variant
    If IsNull(Value) Then
        ExtractDate = Null
        Exit Function
    End If

    ' Using a static, we avoid re-creating the same regex object for every call '
    Static regex As Object
    ' Initialise the Regex object '
    If regex Is Nothing Then
        Set regex = CreateObject("vbscript.regexp")
        With regex
            .Global = True
            .IgnoreCase = True
            .MultiLine = True
            .Pattern = "(\d+\/\d+/\d+\s+\d+:\d+:\d+\s+\w+|\d+-\w+-\d+\s+\d+:\d+:\d+)"
        End With
    End If
    ' Test the value against the pattern '
    Dim matches As Object
    Set matches = regex.Execute(Value)
    If matches.Count > 0 Then
        ' Convert the match to a Date if we can '
        ExtractDate = CDate(matches(0).Value)
    Else
        ' No match found, jsut return Null '
        ExtractDate = Null
    End If
End Function
Public Function GetNumbers(strInput) As String
' Returns the numeric characters within a string in
' sequence in which they are found within the string
    Dim strResult As String, strCh As String
    Dim intI   As Integer
    If Not IsNull(strInput) Then
        For intI = 1 To Len(strInput)
            strCh = Mid(strInput, intI, 1)
            Select Case strCh
            Case "0" To "9"
                strResult = strResult & strCh
            Case Else
            End Select
        Next intI
    End If
    GetNumbers = strResult
End Function

Public Function GetMnRData_Vector(Optional sFile As String, Optional DataDate As String)
    DoCmd.SetWarnings True
    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim SelectedRange, WorkingRange As Range
    Dim i, ii, RecdX, x As Integer
    Dim RateX, MHCValueX As Double
    Dim SQLz, DataX(), DataTypeX, UnitTypeX, UnitSubTypeX, ReadDateX, VinX, UnitNx As String
    Dim lastRow As Long
    Dim TotalRows As Long
    Dim LogActX As String
    Dim ClientIDx, ClientGroupIDx As Long
    If sFile = "" Then sFile = "C:\FAAS\\FleetImports\MandR\MandRSygma_Vector.xlsx"
    If Not IsDate(DataDate) Then DataDate = DateSerial(Year(Date), Month(Date) - 0, 0) Else DataDate = DateSerial(Year(DataDate), Month(DataDate) - 0, 0)
    LogDateX = CDate(DataDate)
    Call MessageUser("Open", "Getting Vector Data", "...this should be quick...")
    'GoTo PostData
    On Error GoTo -1    'FixMe
    On Error GoTo 0
    Set xlApp = CreateObject("Excel.Application")
    If sFile Like "*.xlsx" Then sFile = sFile Else sFile = sFile & ".xlsx"
CheckIfOpen:
    If IsFileOpen(sFile) = True Then
        Set xlApp = GetObject(, "Excel.application")
        xlApp.Application.DisplayAlerts = False
        xlApp.Application.Quit
        Set xlApp = Nothing
    End If
    ' Test for sFile
    If IsNull(sFile) Then sFile = "C:\FAAS\\FleetImports\MandR\MandRSygma_Vector.xlsx"
    'Clear Temp Table
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From MandRData_Temp")
    DoCmd.SetWarnings True
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = False
        Set xlWB = .Workbooks.Open(sFile, , False)
        Call MessageUser("Update", "Collecting Vector MnR Data", "...the original file will not be changed...")
        For x = 1 To 4
            On Error GoTo NoDataOnPage
            .Worksheets(.Worksheets(x).Name).Select
            On Error GoTo FixMe
            lastRow = .Range("C" & .rows.Count).End(xlUp).Row     'How many populated rows in spreadsheet
            '        If LastRow < 3 Then GoTo PostData
            TotalRows = TotalRows + lastRow
            ' LastRow = 5    '- for testing only
            ClientIDx = 3674
            ClientGroupIDx = 88
            .Application.Range("H10").Select      'Test Date
            If CDate(.Application.Selection.Text) < CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX)) Then
                Dim Committ, FileString As String
                Committ = MsgBox("You're collecting data with date ( " & CDate(.Application.Selection.Text) & " )" & _
                                 vbNewLine & "That is before the Activity Date you requested ( " & CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX)) & " ) " & _
                                 vbNewLine & vbNewLine & "Do wish to continue?", vbQuestion + vbYesNo + vbCritical, "Importing Data")
                If Committ = vbNo Then GoTo FixMe
            End If
            Dim Unit"14) As String
            .Application.Columns().Select
            .Application.Selection.UnMerge
            If x > 1 Then .Application.Range("A7").Select Else .Application.Range("A8").Select
            Unit "13) = .Application.Selection.Text"
            Unit"11) = SplitString(Unit"13), " - ", 0)
            Unit"12) = SplitString(Unit"13), "SYGMA Network ", 1)
            Unit"13) = SplitString(Unit"13), "-", 0)

            LogActX = CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX))

            For i = 10 To lastRow
                .Application.Range("C" & i).Select    'UnitNum
                Unit "1) = .Application.Selection.Text"

                .Application.Range("D" & i).Select    'RONumber
                Unit "2) = .Application.Selection.Text"

                .Application.Range("F" & i).Select    'InvoiceNum
                Unit "3) = .Application.Selection.Text"

                .Application.Range("H" & i).Select    'InvoiceDate / RODate
                Unit "4) = .Application.Selection.Text"

                .Application.Range("I" & i).Select    'VRMS
                Unit "5) = Trim(.Application.Selection.Text)"

                .Application.Range("J" & i).Select    'Usage
                Unit "6) = .Application.Selection.Text"

                .Application.Range("K" & i).Select    'PartsCost
                Unit "7) = .Application.Selection.Text"

                .Application.Range("M" & i).Select    'LaborCost
                Unit "8) = .Application.Selection.Text"

                .Application.Range("N" & i).Select    'LaborHours
                Unit "9) = .Application.Selection.Text"

                .Application.Range("O" & i).Select    'OutsideCost
                Unit "10) = .Application.Selection.Text"

                .Application.Range("C" & i).Select                               'Read UnitNumber / VIN
                UnitNx = Trim(.Application.Selection.Text)
                VinX = Nz(DLookup("UnitVin", "Units", "ClientID=" & ClientIDx & " AND UnitUnitNum='" & Unit"1) & "'"), 0)
                UnitIDx = Nz(DLookup("UnitID", "Units", "ClientID=" & ClientIDx & " AND UnitUnitNum='" & Unit"1) & "'"), 0)
                UnitTypeX = Nz(DLookup("UnitType", "Units", "ClientID=" & ClientIDx & " AND UnitUnitNum='" & Unit"1) & "'"), 0)
                UnitSubTypeX = Nz(DLookup("UnitSubType", "Units", "ClientID=" & ClientIDx & " AND UnitUnitNum='" & Unit"1) & "'"), 0)
                If Not IsNumeric(Unit"7)) Then Unit"7) = 0
                If Not IsNumeric(Unit"8)) Then Unit"8) = 0
                If Not IsNumeric(Unit"9)) Then Unit"9) = 0
                If Not IsNumeric(Unit"10)) Then Unit"10) = 0
                SQLz = ""
                SQLz = SQLz & "INSERT INTO MandRData_Temp (  "
                SQLz = SQLz & "ClientGroupID "
                SQLz = SQLz & ", ClientID "
                SQLz = SQLz & ", ActivityMonth "
                SQLz = SQLz & ", UnitID "
                SQLz = SQLz & ", UnitNumber "
                SQLz = SQLz & ", VIN "
                SQLz = SQLz & ", FleetCodeDescr "
                SQLz = SQLz & ", CompanyCodeDescr "
                SQLz = SQLz & ", CompanyDivision "
                SQLz = SQLz & ", TypeDescr "
                SQLz = SQLz & ", VMRS "
                SQLz = SQLz & ", SystemCodeDescr "
                SQLz = SQLz & ", RONumber "
                SQLz = SQLz & ", RODateTime "
                SQLz = SQLz & ", [Usage] "
                SQLz = SQLz & ", LaborHours "
                SQLz = SQLz & ", LaborCost "
                SQLz = SQLz & ", OutsideCost "
                SQLz = SQLz & ", FluidCost "
                SQLz = SQLz & ", PartCost "
                SQLz = SQLz & ", TotalCost "
                SQLz = SQLz & ", LTD ) "
                SQLz = SQLz & "SELECT   "
                SQLz = SQLz & "  " & ClientGroupIDx & " "
                SQLz = SQLz & ", " & ClientIDx & " "
                SQLz = SQLz & ",'" & LogActX & "' "
                SQLz = SQLz & ", " & UnitIDx & " "
                SQLz = SQLz & ",'" & Unit"1) & "' "
                SQLz = SQLz & ",'" & VinX & "' "
                SQLz = SQLz & ",'" & Unit"13) & "' "                   'FleetCodeDescr "
                SQLz = SQLz & ",'SYGMA Network Inc.' "                      'CompanyCodeDescr "
                SQLz = SQLz & ",'" & Unit"12) & "' "                   'CompanyDivision "
                SQLz = SQLz & ",'" & UnitTypeX & " " & UnitSubTypeX & "'"   'TypeDescr "
                SQLz = SQLz & ",'" & Unit"5) & "' "                    'VMRS
                SQLz = SQLz & ", ''"                                        'SystemCodeDescr"
                SQLz = SQLz & ", '" & Unit"3) & "' "                   'RONumber
                SQLz = SQLz & ", '" & Unit"4) & "' "                   'RODateTime"
                SQLz = SQLz & ", " & Unit"6) & " "                     'Usage "
                SQLz = SQLz & ", " & Unit"9) & " "                     'LaborHours "
                SQLz = SQLz & ", " & Unit"8) & " "                     'LaborCost "
                SQLz = SQLz & ", " & Unit"10) & ""                     'OutsideCost "
                SQLz = SQLz & ", 0"                                         'FluidCost "
                SQLz = SQLz & ", " & CCur(Unit"7)) & ""                'PartCost "
                SQLz = SQLz & ", " & CCur(Unit"7)) + CCur(Unit"8)) & ""     'TotalCost "
                SQLz = SQLz & ", 0"                                         'LTD "
                SQLz = SQLz & "; "

                Call MessageUser("Update", "VIN:" & VinX & " ", UnitTypeX & " " & UnitSubTypeX & " Tab-" & x & " " & i - 10 & " of " & lastRow - 10 & "")

                'Insert into the Temp Table data records
                DoCmd.SetWarnings False
                If UnitIDx > 0 Then DoCmd.RunSQL (SQLz)
                DoCmd.SetWarnings True
            Next i
NoDataOnPage:
        Next x
    End With
    xlApp.Application.DisplayAlerts = False
    xlApp.Application.Quit
    Set xlApp = Nothing

PostData:
    Call MessageUser("Update", "Clearing Exisiting Data", "...Only what is to be replaced...")
    SQLz = "DELETE * From MandRData_BETL WHERE ClientID=3674 "
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    'Merge Tempdata with ATLAAS Core Fleet Info
    Call MessageUser("Update", "Pushing To SQL Server", "...Just a bit more...")
    SQLz = ""
    SQLz = SQLz & "INSERT INTO MandRData_BETL ("
    SQLz = SQLz & "  ClientGroupID "
    SQLz = SQLz & ", ClientID "
    SQLz = SQLz & ", ActivityMonth "
    SQLz = SQLz & ", UnitID "
    SQLz = SQLz & ", UnitNumber "
    SQLz = SQLz & ", VIN "
    SQLz = SQLz & ", FleetCodeDescr "
    SQLz = SQLz & ", CompanyCodeDescr "
    SQLz = SQLz & ", CompanyDivision "
    SQLz = SQLz & ", TypeDescr "
    SQLz = SQLz & ", VMRS "
    SQLz = SQLz & ", SystemCodeDescr "
    SQLz = SQLz & ", RONumber "
    SQLz = SQLz & ", RODateTime "
    SQLz = SQLz & ", [Usage] "
    SQLz = SQLz & ", LaborHours "
    SQLz = SQLz & ", LaborCost "
    SQLz = SQLz & ", OutsideCost "
    SQLz = SQLz & ", FluidCost "
    SQLz = SQLz & ", PartCost "
    SQLz = SQLz & ", TotalCost "
    SQLz = SQLz & ", LTD ) "
    SQLz = SQLz & "SELECT ClientGroupID "
    SQLz = SQLz & ", ClientID "
    SQLz = SQLz & ", ActivityMonth "
    SQLz = SQLz & ", UnitID "
    SQLz = SQLz & ", UnitNumber "
    SQLz = SQLz & ", VIN "
    SQLz = SQLz & ", FleetCodeDescr "
    SQLz = SQLz & ", CompanyCodeDescr "
    SQLz = SQLz & ", CompanyDivision "
    SQLz = SQLz & ", TypeDescr "
    SQLz = SQLz & ", VMRS "
    SQLz = SQLz & ", SystemCodeDescr "
    SQLz = SQLz & ", RONumber "
    SQLz = SQLz & ", RODateTime "
    SQLz = SQLz & ", Max(Usage) AS MaxOfUsage "
    SQLz = SQLz & ", Sum(LaborHours) AS SumOfLaborHours "
    SQLz = SQLz & ", Sum(LaborCost) AS SumOfLaborCost "
    SQLz = SQLz & ", Sum(OutsideCost) AS SumOfOutsideCost "
    SQLz = SQLz & ", Sum(FluidCost) AS SumOfFluidCost "
    SQLz = SQLz & ", Sum(PartCost) AS SumOfPartCost "
    SQLz = SQLz & ", Sum(TotalCost) AS SumOfTotalCost "
    SQLz = SQLz & ", Sum(LTD) AS SumOfLTD "
    SQLz = SQLz & "FROM MandRData_Temp "
    SQLz = SQLz & "GROUP BY ClientGroupID, ClientID, ActivityMonth, UnitID, UnitNumber, VIN, FleetCodeDescr, CompanyCodeDescr"
    SQLz = SQLz & ", CompanyDivision, TypeDescr, VMRS, SystemCodeDescr, RONumber, RODateTime; "
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call MessageUser("Close")
    Call SendImportEmailMsg("Sygma-Vector", ClientGroupIDx, LogActX, TotalRows)
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From MandRData_Temp")
    DoCmd.SetWarnings True
    Exit Function
FixMe:
    Call MessageUser("Close")
    xlApp.Application.DisplayAlerts = False
    xlApp.Application.Quit
    MsgBox "This file doesn't have the proper MNR data" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"

End Function
Public Function GetMnRData_Costco(Optional sFile As String, Optional DataDate As String)
'Costco's DataProviderID  is 103

    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim SelectedRange, WorkingRange As Range
    Dim SQLz, DataX(), DataTypeX, UnitTypeX, LeaseTypeX, TermX, ISDx, Poolx, AGMTx, ReadDateX, VinX, UnitNx As String
    Dim FullFile As String
    Dim LogActX As String
    Dim Locz   As String
    'Dim Yearz As String
    Dim lastRow As Long
    Dim TotalRows As Long
    Dim ClientIDx, ClientGroupIDx As Long
    Dim i, RecdX, x, xi As Integer
    Dim RateX, MHCValueX As Double
    'Yearz = Right(DataDate, 4)
    DoCmd.SetWarnings True
    'Datez = SplitString(SplitString(sFile, "Amerit\", 1), "_", 0) & "_" & SplitString(SplitString(sFile, "Amerit\", 1), "_", 1) & "_"
    'Datez = SplitString(SplitString(sFile, "Amerit\", 1), " ", 0)
    'sFile = "\\fa-file\Public\FLEETADVANTAGE\CUSTOMERS\Costco\Maintenance\"
    sFile = "\\fa-file\public\FLEETADVANTAGE\IT\Imports\Costco\MilesHours\"
    If Not IsDate(DataDate) Then DataDate = DateSerial(Year(Date), Month(Date) - 0, 0) Else DataDate = DateSerial(Year(DataDate), Month(DataDate) - 0, 0)
    LogDateX = CDate(DataDate)
    Call MessageUser("Open", "Getting Excel Data", "...this should be quick...")
    For xi = 1 To 6
        If xi = 1 Then Locz = "WPB.xlsx"
        If xi = 2 Then Locz = "Tracy.xlsx"
        If xi = 3 Then Locz = "Tolleson.xlsx"
        If xi = 4 Then Locz = "Sumner.xlsx"
        If xi = 5 Then Locz = "Monroe.xlsx"
        If xi = 6 Then Locz = "Mira Loma.xlsx"
        FullFile = sFile & Locz

        '    GoTo PostData
        On Error GoTo -1    'FixMe
        On Error GoTo 0
        Set xlApp = CreateObject("Excel.Application")
        '    If FullFile Like "*.xlsx" Then FullFile = FullFile Else FullFile = FullFile & ".xlsx"
CheckIfOpen:
        If IsFileOpen(FullFile) = True Then
            Set xlApp = GetObject(, "Excel.application")
            xlApp.Application.DisplayAlerts = False
            xlApp.Application.Quit
            Set xlApp = Nothing
        End If
        'Clear Temp Table
        DoCmd.SetWarnings False
        If xi = 1 Then DoCmd.RunSQL ("Delete * From CostcoDataTemp")
        DoCmd.SetWarnings True
        Set xlApp = CreateObject("Excel.Application")
        With xlApp
            .Visible = False
            Set xlWB = .Workbooks.Open(FullFile, , False)
            Call MessageUser("Update", "Collecting Costco MnR Data", "...the original file will not be changed...")
            For x = 1 To 2
ReadTabs:
                On Error GoTo NoDataOnPage
                .Worksheets(.Worksheets(x).Name).Select
                On Error GoTo FixMe
                ' How many populated rows in spreadsheet
                lastRow = .Range("D" & .rows.Count).End(xlUp).Row
                If lastRow < 3 Then GoTo NoDataOnPage
                ClientIDx = 3410
                ClientGroupIDx = 45

                .Application.Columns().Select
                .Application.Selection.ClearComments

                If Locz Like "*Tolleson*" Then
                    .Application.Range("B1").Select
                    .Application.Selection.EntireColumn.Insert
                End If

                If Locz Like "*Sumner*" Then
                    .Application.Range("B1").Select
                    .Application.Selection.EntireColumn.Delete
                End If

                .Application.Range("D3").Select      'Test Date
                If CDate(.Application.Selection.Text) < CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX)) Then
                    Dim Committ, FileString As String
                    Committ = MsgBox("You're collecting data with date ( " & CDate(.Application.Selection.Text) & " )" & _
                                     vbNewLine & "That is before the Activity Date you requested ( " & CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX)) & " ) " & _
                                     vbNewLine & vbNewLine & "Do wish to continue?", vbQuestion + vbYesNo + vbCritical, "Importing Data")
                    If Committ = vbNo Then GoTo FixMe
                End If
                .Application.Range("C1").Select                                     'Read C1 for Location as Pool
                Poolx = .Application.Selection.Text
                If Poolx = "" Or Len(Poolx) < 10 Then
                    .Application.Range("C2").Select                                     'Read C2 for Location as Pool
                    Poolx = .Application.Selection.Text
                End If
                LogActX = CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX))

                For i = 3 To lastRow
                    .Application.Range("A" & i).Select                               'Read UnitNumber / VIN
                    UnitNx = Trim(.Application.Selection.Text)
                    '  .Application.Range("B" & i).Select
                    'VINx = Trim(.Application.Selection.Text)
                    VinX = Nz(DLookup("UnitVin", "Units", "ClientID=3410 AND UnitUnitNum='" & UnitNx & "'"), 0)
                    SQLz = ""
                    SQLz = SQLz & "INSERT INTO CostcoDataTemp ("
                    SQLz = SQLz & "UnitID, UnitNum, VIN, Pool, AGMT, ReadDateMiles, MilesLTD, MilesRate, ReadDateHours, HoursLTD, "
                    SQLz = SQLz & "HoursRate, StandByDate, StandByRate, ReadDateLiftGate, LiftGateCyclesLTD, LiftGateRate, UID, "
                    SQLz = SQLz & "UnitType, LeaseType, InServDate, LeaseTerm, LogAct ) "
                    SQLz = SQLz & "SELECT "
                    SQLz = SQLz & Nz(DLookup("UnitID", "Units", "ClientID=3410 AND UnitUnitNum='" & UnitNx & "'"), 0) & ", "            'UnitID
                    SQLz = SQLz & "'" & UnitNx & "', "                                                                                  'UnitNum
                    SQLz = SQLz & "'" & VinX & "', "                                                                                    'UnitVIN
                    SQLz = SQLz & "'Costco - " & SplitString(Locz, ".", 0) & "', "                                                      'Pool


                    .Application.Range("C" & i).Select                               'Check If Reefer - Read AGMT
                    If .Application.Selection.Text Like "*Refrigerated*" Or .Application.Selection.Text Like "*Reefer*" Then AGMTx = "Refrigerated" Else AGMTx = "Dry"
                    SQLz = SQLz & "'" & AGMTx & "', "                                'AGMT

                    .Application.Range("D" & i).Select                               'Read Miles Date
                    ReadDateX = Trim(.Application.Selection.Text)
                    SQLz = SQLz & "#" & CDate(ReadDateX) & "#, "

                    .Application.Range("E" & i).Select                               'Read Miles
                    SQLz = SQLz & " " & CDbl(Nz(.Application.Selection.Text, 0)) & ", "   'Miles LTD
                    SQLz = SQLz & "0, "                                              'Read Mile Rate There are no Miles Rate for costco

                    .Application.Range("F" & i).Select                               'Test if there are Hours
                    If AGMTx <> "Dry" Then
                        'If .Application.Selection.Text & "test" <> "test" Then
                        SQLz = SQLz & "#" & CDate(ReadDateX) & "#, "                 'Read Hours Date Same a Miles Date
                        SQLz = SQLz & .Application.Selection.Text & ", "             'Read Hours
                    Else
                        SQLz = SQLz & "#1/1/1900#, "                                 'Read Hours Date Same a Miles Date
                        SQLz = SQLz & "0, "                                          'Read Hours
                    End If
                    SQLz = SQLz & "0, "                                              'Read Hours Rate
                    SQLz = SQLz & "#1/1/1900#, "                                     'Read Stand By Date
                    SQLz = SQLz & "0, "                                              'Read Stand By Rate
                    SQLz = SQLz & "#1/1/1900#, "                                     'Read Lift Gate Start Date
                    SQLz = SQLz & "0, "                                              'Read Lift Gate Cycles
                    SQLz = SQLz & "0, "                                              'Read Lift Gate Rate
                    SQLz = SQLz & "'" & VinX & CLng(CDate(ReadDateX)) & "', "        'Make and get UID
                    SQLz = SQLz & "'Trailer', "                                      'Read UnitType
                    SQLz = SQLz & "'FSL', "                                          'LeaseType
                    SQLz = SQLz & "#" & Nz(DLookup("UnitInserv", "Units", "UnitVIN='" & VinX & "'"), "1/1/1900") & "#, "
                    SQLz = SQLz & "0, "
                    SQLz = SQLz & "#" & CDate(LogActX) & "#"                         'LogDate
                    SQLz = SQLz & ";"

                    Call MessageUser("Update", "File:" & SplitString(Locz, ".", 0) & " (" & xi & " of 6), Tab:" & x & "", AGMTx & " Unit " & i - 2 & " of " & lastRow - 2 & "")

                    'Insert into the Temp Table data records
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL (SQLz)
                    DoCmd.SetWarnings True
                Next i
NoDataOnPage:
            Next x
        End With
        xlApp.Application.DisplayAlerts = False
        xlApp.Application.Quit
        Set xlApp = Nothing
    Next xi
PostData:

    Call MessageUser("Update", "Clearing Exisiting Data", "...Only what is to be replaced...")
    SQLz = "DELETE * From TrailerMilesHours " & _
           "WHERE EXISTS " & _
           "(Select * From CostcoDataTemp Where CostcoDataTemp.LogAct = TrailerMilesHours.LogActivityDate AND CostcoDataTemp.UnitID = TrailerMilesHours.UnitID);"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    'Merge Tempdata with ATLAAS Core Fleet Info
    Call MessageUser("Update", "Pushing To SQL Server", "...Just a bit more...")
    SQLz = ""
    SQLz = "SELECT TrailerMilesHours.UnitID, TrailerMilesHours.LogActivityDate AS PrevLogActivityDate, TrailerMilesHours.MilesLTD AS PrevMilesLTD, Nz([TrailerMilesHours.HoursLTD],0) AS PrevHoursLTD, TrailerMilesHours.ProviderID, Units.ClientID " & _
           "FROM (TrailerMilesHours INNER JOIN " & _
           "(Select Max(LogActivityDate) As LastReportDate, UnitID From TrailerMilesHours Group By UnitID )  AS LastReport " & _
           "ON TrailerMilesHours.UnitID = LastReport.UnitID AND TrailerMilesHours.LogActivityDate = LastReport.LastReportDate) " & _
           "INNER JOIN Units ON TrailerMilesHours.UnitID = Units.UnitID " & _
           "WHERE Units.ClientID  =3410; "
    'MnRData.ProviderID)=103))
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "MnRDataDataLastRead"
    Set qdf = CurrentDb.CreateQueryDef("MnRDataDataLastRead", SQLz)
    On Error GoTo FixMe

    SQLz = ""
    SQLz = SQLz & "INSERT INTO TrailerMilesHours ( UnitID, UnitNumber, VIN, LogActivityDate, ReadDate, Miles, Hours, MilesLTD, CPM, HoursLTD, CPH, ProviderID, MaintLeaseType ) "
    SQLz = SQLz & "SELECT CostcoDataTemp.UnitID, CostcoDataTemp.UnitNum, CostcoDataTemp.VIN, CostcoDataTemp.LogAct, CostcoDataTemp.ReadDateMiles, [MilesLTD]-[PrevMilesLTD] AS Miles, [HoursLTD]-[PrevHoursLTD] AS Hours, "
    SQLz = SQLz & "CostcoDataTemp.MilesLTD, CostcoDataTemp.MilesRate, CostcoDataTemp.HoursLTD, CostcoDataTemp.HoursRate, 103 AS ProviderID, CostcoDataTemp.LeaseType "
    SQLz = SQLz & "FROM CostcoDataTemp LEFT JOIN MnRDataDataLastRead ON CostcoDataTemp.UnitID = MnRDataDataLastRead.UnitID;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    TotalRows = DCount("UnitID", "CostcoDataTemp", "")
    Call TrailerMilesHoursDataValidator(LogActX, 103)
    Call MessageUser("Close")
    Call SendImportEmailMsg("Costco", ClientGroupIDx, LogActX, TotalRows)
    Exit Function
FixMe:
    Call MessageUser("Close")
    xlApp.Application.DisplayAlerts = False
    xlApp.Application.Quit
    MsgBox "File " & Locz & " doesn't have the proper MNR data" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"
End Function
Public Function GetMnRData_SygmaInternal(Optional sFile As String, Optional DataDate As String)
'Costco's DataProviderID  is 103
    DoCmd.SetWarnings True
    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim SelectedRange, WorkingRange As Range
    'Dim WSheet As Worksheet

    Dim i, RecdX, x As Integer
    Dim RateX, MHCValueX As Double
    Dim SQLz, DataX(), DataTypeX, UnitTypeX, LeaseTypeX, TermX, ISDx, Poolx, AGMTx, ReadDateX, VinX, UnitNx As String
    Dim lastRow As Long
    Dim TotalRows As Long
    Dim LogActX As String
    Dim ClientIDx, ClientGroupIDx As Long
    If Not IsDate(DataDate) Then DataDate = DateSerial(Year(Date), Month(Date) - 0, 0) Else DataDate = DateSerial(Year(DataDate), Month(DataDate) - 0, 0)
    LogDateX = CDate(DataDate)
    Call MessageUser("Open", "Getting Excel Data", "...this should be quick...")
    '    GoTo PostData
    On Error GoTo -1    'FixMe
    On Error GoTo 0
    Set xlApp = CreateObject("Excel.Application")
    If sFile Like "*.xlsx" Then sFile = sFile Else sFile = sFile & ".xlsx"
CheckIfOpen:
    If IsFileOpen(sFile) = True Then
        Set xlApp = GetObject(, "Excel.application")
        xlApp.Application.DisplayAlerts = False
        xlApp.Application.Quit
        Set xlApp = Nothing
    End If
    ' Test for sFile
    'If IsNull(sFile) Then sFile = "C:\FAAS\\FleetImports\MnR_Penske.xlsx"
    'Clear Temp Table
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From CostcoDataTemp")
    DoCmd.SetWarnings True
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = False
        Set xlWB = .Workbooks.Open(sFile, , False)
        Call MessageUser("Update", "Collecting Costco MnR Data", "...the original file will not be changed...")
        For x = 1 To 2
            On Error GoTo NoDataOnPage
            .Worksheets(.Worksheets(x).Name).Select
            On Error GoTo FixMe
            ' How many populated rows in spreadsheet
            lastRow = .Range("D" & .rows.Count).End(xlUp).Row
            '        If LastRow < 3 Then GoTo PostData
            TotalRows = TotalRows + lastRow
            ' ReDim DataX(1 To LastRow, 1 To 19)
            ' LastRow = 5    '- for testing only
            ClientIDx = 3410
            ClientGroupIDx = 45
            .Application.Range("D3").Select      'Test Date

            If CDate(.Application.Selection.Text) < CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX)) Then
                Dim Committ, FileString As String
                Committ = MsgBox("You're collecting data with date ( " & CDate(.Application.Selection.Text) & " )" & _
                                 vbNewLine & "That is before the Activity Date you requested ( " & CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX)) & " ) " & _
                                 vbNewLine & vbNewLine & "Do wish to continue?", vbQuestion + vbYesNo + vbCritical, "Importing Data")
                If Committ = vbNo Then GoTo FixMe
            End If
            .Application.Range("C1").Select                                     'Read C1 for Location as Pool
            Poolx = .Application.Selection.Text
            If Poolx = "" Or Len(Poolx) < 10 Then
                .Application.Range("C2").Select                                     'Read C2 for Location as Pool
                Poolx = .Application.Selection.Text
            End If
            LogActX = CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX))

            For i = 3 To lastRow
                .Application.Range("A" & i).Select                               'Read UnitNumber / VIN
                UnitNx = Trim(.Application.Selection.Text)
                '  .Application.Range("B" & i).Select
                'VINx = Trim(.Application.Selection.Text)
                VinX = Nz(DLookup("UnitVin", "Units", "ClientID=3410 AND UnitUnitNum='" & UnitNx & "'"), 0)
                SQLz = ""
                SQLz = SQLz & "INSERT INTO CostcoDataTemp ("
                SQLz = SQLz & "UnitID, UnitNum, VIN, Pool, AGMT, ReadDateMiles, MilesLTD, MilesRate, ReadDateHours, HoursLTD, "
                SQLz = SQLz & "HoursRate, StandByDate, StandByRate, ReadDateLiftGate, LiftGateCyclesLTD, LiftGateRate, UID, "
                SQLz = SQLz & "UnitType, LeaseType, InServDate, LeaseTerm, LogAct ) "
                SQLz = SQLz & "SELECT "
                SQLz = SQLz & Nz(DLookup("UnitID", "Units", "ClientID=3410 AND UnitUnitNum='" & UnitNx & "'"), 0) & ", "            'UnitID
                SQLz = SQLz & "'" & UnitNx & "', "                                                                                  'UnitNum
                SQLz = SQLz & "'" & VinX & "', "                                                                                    'UnitVIN
                SQLz = SQLz & "'" & Poolx & "', "                                                                                   'Pool
                'Sqlz = Sqlz & "'" & Nz(DLookup("UnitUnitNum", "Units", "UnitVIN='" & VINx & "'"), Trim(.Application.Selection.Text)) & "', "  'UnitNum
                'Sqlz = Sqlz & "'" & VINx & "', "                                                              'VIN


                .Application.Range("C" & i).Select                               'Check If Reefer - Read AGMT
                If .Application.Selection.Text Like "*Refrigerated*" Or .Application.Selection.Text Like "*Reefer*" Then AGMTx = "Reefrigerated" Else AGMTx = "Dry"
                SQLz = SQLz & "'" & AGMTx & "', "                                'AGMT

                .Application.Range("D" & i).Select                               'Read Miles Date
                ReadDateX = Trim(.Application.Selection.Text)
                SQLz = SQLz & "#" & CDate(ReadDateX) & "#, "

                .Application.Range("E" & i).Select                               'Read Miles
                SQLz = SQLz & " " & CDbl(Nz(.Application.Selection.Text, 0)) & ", "   'Miles LTD
                SQLz = SQLz & "0, "                                              'Read Mile Rate There are no Miles Rate for costco

                .Application.Range("F" & i).Select                               'Test if there are Hours
                If .Application.Selection.Text & "test" <> "test" Then
                    SQLz = SQLz & "#" & CDate(ReadDateX) & "#, "                 'Read Hours Date Same a Miles Date
                    SQLz = SQLz & .Application.Selection.Text & ", "             'Read Hours
                Else
                    SQLz = SQLz & "#1/1/1900#, "                                 'Read Hours Date Same a Miles Date
                    SQLz = SQLz & "0, "                                          'Read Hours
                End If
                SQLz = SQLz & "0, "                                              'Read Hours Rate
                SQLz = SQLz & "#1/1/1900#, "                                     'Read Stand By Date
                SQLz = SQLz & "0, "                                              'Read Stand By Rate
                SQLz = SQLz & "#1/1/1900#, "                                     'Read Lift Gate Start Date
                SQLz = SQLz & "0, "                                              'Read Lift Gate Cycles
                SQLz = SQLz & "0, "                                              'Read Lift Gate Rate
                SQLz = SQLz & "'" & VinX & CLng(CDate(ReadDateX)) & "', "        'Make and get UID
                SQLz = SQLz & "'Trailer', "                                      'Read UnitType
                SQLz = SQLz & "'FSL', "                                          'LeaseType
                SQLz = SQLz & "#" & Nz(DLookup("UnitInserv", "Units", "UnitVIN='" & VinX & "'"), "1/1/1900") & "#, "
                SQLz = SQLz & "0, "
                SQLz = SQLz & "#" & CDate(LogActX) & "#"                         'LogDate
                SQLz = SQLz & ";"
                If x = 1 Then
                    Call MessageUser("Update", "VIN:" & VinX & "", "Dry Van Unit " & i - 2 & " of " & lastRow - 2 & "")
                Else
                    Call MessageUser("Update", "VIN:" & VinX & "", "Reefer Unit " & i - 2 & " of " & lastRow - 2 & "")
                End If
                'Insert into the Temp Table data records
                DoCmd.SetWarnings False
                DoCmd.RunSQL (SQLz)
                DoCmd.SetWarnings True
            Next i
NoDataOnPage:
        Next x
    End With
    xlApp.Application.DisplayAlerts = False
    xlApp.Application.Quit
    Set xlApp = Nothing

PostData:

    Call MessageUser("Update", "Clearing Exisiting Data", "...Only what is to be replaced...")
    SQLz = "DELETE * From TrailerMilesHours " & _
           "WHERE EXISTS " & _
           "(Select * From CostcoDataTemp Where CostcoDataTemp.LogAct = TrailerMilesHours.LogActivityDate AND CostcoDataTemp.UnitID = TrailerMilesHours.UnitID);"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    'Merge Tempdata with ATLAAS Core Fleet Info
    Call MessageUser("Update", "Pushing To SQL Server", "...Just a bit more...")
    SQLz = ""
    SQLz = "SELECT TrailerMilesHours.UnitID, TrailerMilesHours.LogActivityDate AS PrevLogActivityDate, TrailerMilesHours.MilesLTD AS PrevMilesLTD, Nz([TrailerMilesHours.HoursLTD],0) AS PrevHoursLTD, TrailerMilesHours.ProviderID, Units.ClientID " & _
           "FROM (TrailerMilesHours INNER JOIN " & _
           "(Select Max(LogActivityDate) As LastReportDate, UnitID From TrailerMilesHours Group By UnitID )  AS LastReport " & _
           "ON TrailerMilesHours.UnitID = LastReport.UnitID AND TrailerMilesHours.LogActivityDate = LastReport.LastReportDate) " & _
           "INNER JOIN Units ON TrailerMilesHours.UnitID = Units.UnitID " & _
           "WHERE Units.ClientID  =3410; "
    'MnRData.ProviderID)=103))
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "MnRDataDataLastRead"
    Set qdf = CurrentDb.CreateQueryDef("MnRDataDataLastRead", SQLz)
    On Error GoTo FixMe

    SQLz = ""
    SQLz = SQLz & "INSERT INTO TrailerMilesHours ( UnitID, UnitNumber, VIN, LogActivityDate, ReadDate, Miles, Hours, MilesLTD, CPM, HoursLTD, CPH, ProviderID, MaintLeaseType ) "
    SQLz = SQLz & "SELECT CostcoDataTemp.UnitID, CostcoDataTemp.UnitNum, CostcoDataTemp.VIN, CostcoDataTemp.LogAct, CostcoDataTemp.ReadDateMiles, [MilesLTD]-[PrevMilesLTD] AS Miles, [HoursLTD]-[PrevHoursLTD] AS Hours, "
    SQLz = SQLz & "CostcoDataTemp.MilesLTD, CostcoDataTemp.MilesRate, CostcoDataTemp.HoursLTD, CostcoDataTemp.HoursRate, 103 AS ProviderID, CostcoDataTemp.LeaseType "
    SQLz = SQLz & "FROM CostcoDataTemp LEFT JOIN MnRDataDataLastRead ON CostcoDataTemp.UnitID = MnRDataDataLastRead.UnitID;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call TrailerMilesHoursDataValidator(LogActX, 103)
    Call MessageUser("Close")
    Call SendImportEmailMsg("Costco", ClientGroupIDx, LogActX, TotalRows)
    Exit Function
FixMe:
    Call MessageUser("Close")
    xlApp.Application.DisplayAlerts = False
    xlApp.Application.Quit
    MsgBox "This file doesn't have the proper MNR data" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"
End Function
Function IsFileOpen(fileName As String)
    Dim FileNum As Integer, errnum As Integer

    On Error Resume Next    ' Turn error checking off.
    FileNum = FreeFile()    ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open fileName For Input Lock Read As #FileNum
    Close FileNum           ' Close the file.
    errnum = Err            ' Save the error number that occurred.
    On Error GoTo 0         ' Turn error checking back on.

    Select Case errnum      ' Check to see which error occurred.
    Case 0                  ' No error occurred. File is NOT already open by another user.
        IsFileOpen = False  ' Error number for "Permission Denied. File is already opened by another user.
    Case 70
        IsFileOpen = True   ' Another error occurred.
    Case Else
        Error errnum
    End Select
End Function

Public Function fileName(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        fileName = fileName(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
Function OpenMnRDataFileDialog(ProviderType As String, Optional PassDate As String)
    Dim f      As Object
    Set f = Application.FileDialog(3)
    f.InitialFileName = "P:\FLEETADVANTAGE\CUSTOMERS"
    If ProviderType = "AIM" Then f.InitialFileName = "P:\FLEETADVANTAGE\CUSTOMERS\Maines\AIM"
    If ProviderType = "Penske" Then f.InitialFileName = "P:\FLEETADVANTAGE\CUSTOMERS\Maines\Penske"
    If ProviderType = "Costco" Then f.InitialFileName = "\\fa-file\public\FLEETADVANTAGE\IT\Imports\Costco\MilesHours\"
    f.AllowMultiSelect = False
    f.Filters.Clear
    f.Filters.Add "Excel", "*.xlsx"
    f.Filters.Add "Excel 97", "*.xls"
    f.Filters.Add "CSV", "*.csv"
    f.Filters.Add "All Files", "*.*"
    f.Show
    On Error GoTo FixMe
    Dim Committ, FileString As String
    FileString = f.SelectedItems(1)
    Committ = MsgBox("You're Collecting Data For: " & ProviderType & _
                     vbNewLine & "File: " & FileNameFromPath(FileString) & _
                     vbNewLine & "LogActivity Date: " & DateAdd("m", -1, CDate(DatePart("m", PassDate) & "/1/" & DatePart("yyyy", PassDate))) & _
                     vbNewLine & vbNewLine & "Do wish to continue?", vbQuestion + vbYesNo + vbCritical, "Importing Data")
    If Committ = vbNo Then Exit Function
    ' MsgBox "file choosen = " & f.SelectedItems(1)
    If ProviderType = "AIM" Then Call GetMnRData_AIM(f.SelectedItems(1), PassDate)
    If ProviderType = "Penske" Then Call GetMnRData_Penske(f.SelectedItems(1), PassDate)
    If ProviderType = "Costco" Then Call GetMnRData_Costco(f.SelectedItems(1), PassDate)
    If ProviderType = "SygmaInternal" Then Call GetMnRData_SygmaInternal(f.SelectedItems(1), PassDate)
    Exit Function
FixMe:
    MsgBox "Import Aborted", vbInformation, "ATLAAS Core"
End Function
Function FileNameFromPath(strFullPath As String) As String
    FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function
Public Function GetMnRData_Penske(Optional sFile As String, Optional DataDate As String)
'Penske's DataProviderID  is 102
    DoCmd.SetWarnings True
    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim SelectedRange, WorkingRange As Range
    Dim i, RecdX, x As Integer
    Dim RateX, MHCValueX, MonthlyFeeX As Double
    Dim TargetX, CurrVin, PrevVin, SQLz, DataX(), DataTypeX, UnitTypeX, LeaseTypeX, TermX, ISDx As String
    Dim lastRow As Long
    If Not IsDate(DataDate) Then DataDate = DateSerial(Year(Date), Month(Date) - 0, 0) Else DataDate = DateSerial(Year(DataDate), Month(DataDate) - 0, 0)
    LogDateX = CDate(DataDate)
    Call MessageUser("Open", "Getting Excel Data", "...this should be quick...")
    '    GoTo PostData
    On Error GoTo 0    'FixMe
    Set xlApp = CreateObject("Excel.Application")
    If sFile Like "*.xlsx" Then sFile = sFile Else sFile = sFile & ".xlsx"
CheckIfOpen:
    If IsFileOpen(sFile) = True Then
        Set xlApp = GetObject(, "Excel.application")
        xlApp.Application.DisplayAlerts = False
        xlApp.Application.Quit
        Set xlApp = Nothing
    End If
    ' Test for sFile
    'If IsNull(sFile) Then sFile = "C:\FAAS\\FleetImports\MnR_Penske.xlsx"

    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = True

        Set xlWB = .Workbooks.Open(sFile, , False)
        Call MessageUser("Update", "Collecting Penske MnR Data", "...the original file will not be changed...")
        ' How many populated rows in spreadsheet
        lastRow = .Range("D" & .rows.Count).End(xlUp).Row
        'LastRow = 12 - for testing only
        'Remove Terminated Units
        Call MessageUser("Update", "Removing Terminated Unit", "Why are they there in the 1st place?")
        With .ActiveSheet
            ' .Application.Range("R1").AutoFilter
            '        .Application.Range("R1").AutoFilter Field:=18, Criteria1:="TERMED"
            '        .Application.Range("A2:S" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            '        .ShowAllData
            .Application.Range("N1").AutoFilter Field:=14, Criteria1:=">0"
            .Application.Range("O1").AutoFilter Field:=15, Criteria1:="0"
            .Application.Range("A2:S" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .ShowAllData
            lastRow = .Range("D" & .rows.Count).End(xlUp).Row
            ' Sort the rows based on the VIN in column J
            '.Application.Range("A1").Select
            '.Application.Columns("A:S").Sort key1:=Range("J2"), order1:=xlAscending, Header:=xlYes
        End With
        ReDim DataX(1 To lastRow, 1 To 20)
        ' LastRow = 90    '- for testing only
        For i = 2 To lastRow
            .Application.Range("J" & i).Select                      'Get This Row's VIN
            CurrVin = Trim(.Application.Selection.Text)
            If i < 2 Then Prev = Null
            If i < 3 Then RecdX = i
            If PrevVin <> CurrVin And i > 2 Then RecdX = RecdX + 1
            .Application.Range("L" & i).Select                      'Read DataType
            DataTypeX = Trim(.Application.Selection.Text)
            .Application.Range("K" & i).Select                      'Read UnitType
            UnitTypeX = Trim(.Application.Selection.Text)
            .Application.Range("M" & i).Select                      'Read LeaseType
            LeaseTypeX = Trim(.Application.Selection.Text)
            .Application.Range("H" & i).Select                      'Read In Service Date
            ISDx = Trim(.Application.Selection.Text)
            .Application.Range("I" & i).Select                      'Read Lease Term
            TermX = Trim(.Application.Selection.Text)
            .Application.Range("N" & i).Select                      'Read Monthly Fee
            MonthlyFeeX = CDbl(Nz(.Application.Selection.Text, 0))
            .Application.Range("O" & i).Select                      'Read Rate
            RateX = CDec(Nz(.Application.Selection.Text, 0))
            .Application.Range("S" & i).Select                      'Read Mile/Hours/Cycles LTD Value
            MHCValueX = CDbl(Nz(.Application.Selection.Text, 0))

            DataX(RecdX, 1) = CurrVin                                               'Set VIN
            .Application.Range("M" & i).Select                                      'Read Set Pool
            DataX(RecdX, 2) = Trim(.Application.Selection.Text)
            .Application.Range("Q" & i).Select                                      'Read Set AGMT
            DataX(RecdX, 3) = Trim(.Application.Selection.Text)
            DataX(RecdX, 4) = LogDateX                                              'Set Miles Date
            If DataTypeX = "T" Or DataTypeX = "P" Then DataX(RecdX, 5) = MHCValueX  'Set Miles LTD
            If DataTypeX = "T" Or DataTypeX = "P" Then DataX(RecdX, 6) = RateX      'Set Miles Rate
            DataX(RecdX, 20) = MonthlyFeeX                                          'Set Monthly Fee
            DataX(RecdX, 7) = LogDateX                                              'Set Hours Date
            If DataTypeX = "R" Then DataX(RecdX, 8) = MHCValueX                     'Set Hours LTD
            If DataTypeX = "R" Then
                DataX(RecdX, 9) = RateX                                             'Set Hours Rate
            End If
            DataX(RecdX, 10) = LogDateX                                             'Set Standby Date
            If DataTypeX = "S" Then DataX(RecdX, 11) = RateX                        'Set Stand By Rate
            DataX(RecdX, 12) = LogDateX                                             'Set Liftgate Read Date
            DataX(RecdX, 15) = CurrVin & CLng(CDate(LogDateX))                      'Set UID
            DataX(RecdX, 16) = UnitTypeX                                            'Set UnitType
            DataX(RecdX, 17) = LeaseTypeX                                           'Set Lease Type
            DataX(RecdX, 18) = ISDx                                                 'Set In Service Date
            DataX(RecdX, 19) = TermX                                                'Set Lease Term
            PrevVin = CurrVin
            Call MessageUser("Update", "VIN " & CurrVin & "-" & DataTypeX & " Data", "On Row:" & i & " (Unit " & RecdX & ") of " & lastRow & " data rows")
        Next i
    End With
    xlApp.Application.DisplayAlerts = False
    xlApp.Application.Quit
    Set xlApp = Nothing
    Call MessageUser("Update", "Posting To Temp Tables", "...almost done...")

    'Loop through the DataX array, to read how many actual VINs - Set LastRow
    x = 1
    For i = 2 To lastRow
        If Trim(DataX(i, 1)) & "test" <> "test" Then x = x + 1
    Next i
    lastRow = x

    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From PenskeDataTemp")
    'Insert into the Temp Table data records
    SQLz = ""
    Dim LogActX As String
    Dim ClientIDx, ClientGroupIDx As Long
    LogActX = CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX))
    For i = 2 To lastRow
        SQLz = ""
        SQLz = SQLz & "INSERT INTO PenskeDataTemp ("
        SQLz = SQLz & "UnitNumber, VIN, Pool, AGMT, ReadDateMiles, MilesLTD, MilesRate, ReadDateHours, HoursLTD, "
        SQLz = SQLz & "HoursRate, StandByDate, StandByRate, ReadDateLiftGate, LiftGateCyclesLTD, LiftGateRate, UID, "
        SQLz = SQLz & "UnitType, LeaseType, InServDate, LeaseTerm, LogAct, MonthlyFee ) "
        SQLz = SQLz & "SELECT '"
        SQLz = SQLz & Nz(DLookup("UnitUnitNum", "Units", "UnitVIN='" & DataX(i, 1) & "'"), "") & "', '"    'UnitNumber
        SQLz = SQLz & DataX(i, 1) & "', '"                      ' Vin
        SQLz = SQLz & DataX(i, 2) & "', '"                      ' Pool
        SQLz = SQLz & DataX(i, 3) & "', '"                      ' AGMT
        SQLz = SQLz & CDate(DataX(i, 4)) & "', "                ' Miles Date
        SQLz = SQLz & CDbl(DataX(i, 5)) & ", "                  ' Miles LTD
        SQLz = SQLz & CDbl(DataX(i, 6)) & ", '"                 ' Miles Rate
        SQLz = SQLz & CDate(DataX(i, 7)) & "', "                ' Hours Date
        SQLz = SQLz & CDbl(Nz(DataX(i, 8), 0)) & ", "           ' Hours LTD
        SQLz = SQLz & CDec(Nz(DataX(i, 9), 0)) & ", '"          ' Hours Rate
        SQLz = SQLz & CDate(DataX(i, 10)) & "', "               ' StandBy Date
        SQLz = SQLz & CDbl(Nz(DataX(i, 11), 0)) & ", '"         ' StandBy Rate
        SQLz = SQLz & CDate(DataX(i, 12)) & "', "               ' LiftGate Read Date
        SQLz = SQLz & CDbl(Nz(DataX(i, 13), 0)) & ", "          ' LiftGate Cycles LTD
        SQLz = SQLz & CDbl(Nz(DataX(i, 14), 0)) & ", '"         ' LiftGate Rate
        SQLz = SQLz & DataX(i, 15) & "', '"                     ' UID
        SQLz = SQLz & DataX(i, 16) & "', '"                     ' UnitType
        SQLz = SQLz & DataX(i, 17) & "', #"                     ' LeaseType
        SQLz = SQLz & CDate(Nz(DataX(i, 18), "1/1/1900")) & "#, "   ' InServDate
        SQLz = SQLz & CDbl(Nz(DataX(i, 19), 0)) & ", #"         ' LeaseTerm
        SQLz = SQLz & CDate(LogActX) & "#, "                    ' LogDate
        SQLz = SQLz & CDbl(Nz(DataX(i, 20), 0)) & " "           ' Monthly Fee
        SQLz = SQLz & ";"
        DoCmd.RunSQL (SQLz)
    Next i
    DoCmd.SetWarnings True
    Set xlApp = Nothing
    Call MessageUser("Update", "Collecting Details", "...getting there...")
    i = 2
    Do While (True)
        ClientIDx = Nz(DMax("ClientID", "Units", "UnitVin='" & DataX(i, 1) & "'"), 0)
        ClientGroupIDx = Nz(DMax("ClientGroupID", "Units", "UnitVin='" & DataX(i, 1) & "'"), 0)
        If ClientIDx > 0 Then Exit Do
        i = i + 1
    Loop

PostData:
    Call MessageUser("Update", "Clearing Exisiting Data", "...Only what is to be replaced...")
    SQLz = "DELETE * From MnRData " & _
           "WHERE EXISTS " & _
           "(Select * From PenskeDataTemp Where PenskeDataTemp.LogAct = MnRData.LogActivityDate AND PenskeDataTemp.VIN = MnRData.VIN);"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call MessageUser("Update", "Pushing To SQL Server", "...Just a bit more...")

    'Merge Tempdata with ATLAAS Core Fleet Info
    SQLz = ""
    SQLz = "SELECT UnitID, VIN, LogActivityDate AS PrevLogActivityDate, MilesLTD AS PrevMilesLTD, HoursLTD AS PrevHoursLTD, LiftGateCyclesLTD AS PrevLiftGateLTD, ProviderID " & _
           "From MnRData " & _
           "WHERE (((MnRData.LogActivityDate)=#" & DateAdd("M", -1, LogActX) & "#) AND ((MnRData.ProviderID)=102));"
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "MnRDataDataLastRead"
    Set qdf = CurrentDb.CreateQueryDef("MnRDataDataLastRead", SQLz)
    On Error GoTo FixMe

    SQLz = ""
    SQLz = SQLz & "INSERT INTO MnRData ( UnitID, UnitNumber, VIN, POOL, AGMT, MonthlyFee, ReadDateMiles, PrevMilesLTD, Miles, MilesLTD, MilesRate, "
    SQLz = SQLz & "ReadDateHours, PrevHoursLTD, Hours, HoursLTD, HoursRate, StandByDate, StandByRate, ReadDateLiftGate, "
    SQLz = SQLz & "PrevLiftGateLTD, LiftGateCycles, LiftGateCyclesLTD, LiftGateRate, UID, LogActivityDate, UnitType, ClientID, "
    SQLz = SQLz & "ClientGroupID, LiftGateStartDate, ProviderID ) "
    SQLz = SQLz & "SELECT DISTINCT "
    SQLz = SQLz & "Nz([Units.UnitID],0) AS UnitID, "
    SQLz = SQLz & "PenskeDataTemp.UnitNumber, "
    SQLz = SQLz & "PenskeDataTemp.VIN, "
    SQLz = SQLz & "PenskeDataTemp.POOL, "
    SQLz = SQLz & "PenskeDataTemp.AGMT, "
    SQLz = SQLz & "PenskeDataTemp.MonthlyFee, "
    SQLz = SQLz & "PenskeDataTemp.ReadDateMiles, "
    SQLz = SQLz & "MnRDataDataLastRead.PrevMilesLTD, "
    SQLz = SQLz & "Nz([MilesLTD]-[PrevMilesLTD],0) AS Miles, "
    SQLz = SQLz & "PenskeDataTemp.MilesLTD, PenskeDataTemp.MilesRate, PenskeDataTemp.ReadDateHours, "
    SQLz = SQLz & "MnRDataDataLastRead.PrevHoursLTD, "
    SQLz = SQLz & "Nz([PrevHoursLTD]-[PrevLiftGateLTD],0) AS Hours, "
    SQLz = SQLz & "PenskeDataTemp.HoursLTD, PenskeDataTemp.HoursRate, PenskeDataTemp.StandByDate, PenskeDataTemp.StandByRate, PenskeDataTemp.ReadDateLiftGate, "
    SQLz = SQLz & "MnRDataDataLastRead.PrevLiftGateLTD, "
    SQLz = SQLz & "Nz([LiftGateCyclesLTD]-[PrevLiftGateLTD],0) AS LiftGateCycles, "
    SQLz = SQLz & "PenskeDataTemp.LiftGateCyclesLTD, PenskeDataTemp.LiftGateRate, PenskeDataTemp.UID, "
    SQLz = SQLz & "#" & LogActX & "# AS LogActivityDate, "
    SQLz = SQLz & "PenskeDataTemp.UnitType, " & ClientIDx & ", " & ClientGroupIDx & ", PenskeDataTemp.InServDate, 102 "
    SQLz = SQLz & "FROM (PenskeDataTemp LEFT JOIN Units ON PenskeDataTemp.VIN = Units.UnitVIN) "
    SQLz = SQLz & "LEFT JOIN MnRDataDataLastRead ON MnRDataDataLastRead.VIN = PenskeDataTemp.VIN;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Call PostTrailerMilesHours(LogActX, 102)
    Call TrailerMilesHoursDataValidator(LogActX, 102)
    Call SendImportEmailMsg("Penske", ClientGroupIDx, LogActX, lastRow)
    Exit Function
FixMe:
    Call MessageUser("Close")
    xlApp.Application.Quit
    MsgBox "This file doesn't have the proper MNR data" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"
End Function
Public Function GetMnRData_AIM(sFile As String, Optional DataDate As String)
'AIM's DataProviderID  is 101
    DoCmd.SetWarnings False
    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim SelectedRange, WorkingRange As Range
    Dim i, x   As Integer
    Dim TargetX, SQLz, LogDateX, DataX(), UnitTypeX As String
    Dim lastRow As Long
    Dim TotalRows As Long
    If Not IsDate(DataDate) Then DataDate = DateSerial(Year(Date), Month(Date) - 0, 0) Else DataDate = DateSerial(Year(DataDate), Month(DataDate) - 0, 0)
    LogDateX = CDate(DataDate)
    Call MessageUser("Open", "Getting Excel Data", "...this should be quick...")
    'On Error GoTo FixMe
    Set xlApp = CreateObject("Excel.Application")
    'mFile = sFile
    If sFile Like "*.xlsx" Or sFile Like "*.xls" Then sFile = sFile Else sFile = sFile & ".xlsx"
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = True
        Set xlWB = .Workbooks.Open(sFile, , False)
        Call MessageUser("Update", "Removing Extra cols from Excel", "...the original file will not be changed...")
        With .ActiveSheet
            ' Delete FirstRow
            .Application.rows("1:1").Select
            .Application.Selection.EntireRow.Delete
            lastRow = .Range("D" & .rows.Count).End(xlUp).Row
            ' Format, delete not needed columns
            .Application.Range("R2:R" & lastRow).SpecialCells(xlCellTypeVisible).EntireColumn.Delete
            .Application.Range("P2:P" & lastRow).SpecialCells(xlCellTypeVisible).EntireColumn.Delete
            .Application.Range("N2:N" & lastRow).SpecialCells(xlCellTypeVisible).EntireColumn.Delete
            .Application.Range("L2:L" & lastRow).SpecialCells(xlCellTypeVisible).EntireColumn.Delete
            .Application.Range("J2:J" & lastRow).SpecialCells(xlCellTypeVisible).EntireColumn.Delete
            .Application.Range("H2:H" & lastRow).SpecialCells(xlCellTypeVisible).EntireColumn.Delete
            .Application.Range("G2:G" & lastRow).SpecialCells(xlCellTypeVisible).EntireColumn.Delete
            .Application.Range("F2:F" & lastRow).SpecialCells(xlCellTypeVisible).EntireColumn.Delete
            'Un-Merge columns
            Call MessageUser("Update", "Un-Merging Columns in Excel", "...the original file will not be changed...")
            .Application.Columns(11).Select
            .Application.Selection.UnMerge

            'Delete empty columns
            .Application.Columns(12).Select
            .Application.Selection.EntireColumn.Delete
            'Remove Duplicates
            .Range("A1:R" & lastRow).RemoveDuplicates
            'Insert UID column
            .Application.Range("T1") = "UID"
            .Application.Cells.Select

            ' .Application.Selection.ClearFormats
            .Application.Selection.Font.Name = "Calibri"
            .Application.Selection.Font.Size = 10
            .Application.rows("1:1").Select

            .Application.Selection.Font.Bold = True
            .Application.Selection.Font.Size = 10
            .Application.Cells.Select
            .Application.Selection.RowHeight = 12.75

            .Application.Selection.Columns.AutoFit
            .Application.Range("A2").Select
            .Application.ActiveWindow.FreezePanes = True
            .Application.Range("A1").Select
            .Application.Selection.AutoFilter
            .Application.Selection.Columns.AutoFit

            Call MessageUser("Update", "Removing Non-Trailers", "...the Rentals, Powered, etc...")
            If StartRow > 1 Then lastRow = lastRow + (StartRow - 1)
            .Application.Range("F1").AutoFilter
            .Application.Range("F1").AutoFilter Field:=6, Criteria1:="RENTAL"
            .Application.Range("A2:K" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .ShowAllData
            .Application.Range("F1").AutoFilter
            .Application.Range("F1").AutoFilter Field:=6, Criteria1:="OUTREN"
            .Application.Range("A2:K" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .ShowAllData
            'convert columns to Short dates
            .Application.Columns(8).Select
            .Application.Selection.NumberFormat = "m/d/yyyy"
            .Application.Columns(11).Select
            .Application.Selection.NumberFormat = "m/d/yyyy"
            .Application.Columns(14).Select
            .Application.Selection.NumberFormat = "m/d/yyyy"
            .Application.Columns(17).Select
            .Application.Selection.NumberFormat = "m/d/yyyy"

            'Remove OLD previous month (already reported)
            .Application.Range("H1").AutoFilter
            .Application.Range("H1").AutoFilter Field:=8, Criteria1:="<" & (DateSerial(Year(LogDateX), Month(LogDateX) - 0, 0)) & ""
            .Application.Range("A2:K" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .ShowAllData

            'We're now Keeping Non-Trailer Rows 11/1/2016
            '.Application.Range("C1").AutoFilter
            '.Application.Range("C1").AutoFilter Field:=3, Criteria1:="FRGH"
            '.Application.Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            '.ShowAllData
            '.Application.Range("C1").AutoFilter
            '.Application.Range("C1").AutoFilter Field:=3, Criteria1:="MACK"
            '.Application.Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            '.ShowAllData
            '.Application.Range("C1").AutoFilter
            '.Application.Range("C1").AutoFilter Field:=3, Criteria1:="NVSTR"
            '.Application.Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            '.ShowAllData
            '.Application.Range("C1").AutoFilter
            '.Application.Range("C1").AutoFilter Field:=3, Criteria1:="VOLV"
            '.Application.Range("A2:A" & LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            '.ShowAllData
            lastRow = .Range("D" & .rows.Count).End(xlUp).Row
            .Range("A1:S" & lastRow).RemoveDuplicates

            lastRow = .Range("D" & .rows.Count).End(xlUp).Row
        End With    ' end with ActiveSheet - format, delete columns
        TotalRows = TotalRows + lastRow

        .Application.Range("H2").Select      'Test Date
        If CDate(.Application.Selection.Text) < CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX)) Then
            Dim Committ, FileString As String
            Committ = MsgBox("You're collecting data with date ( " & CDate(.Application.Selection.Text) & " )" & _
                             vbNewLine & "That is before the Activity Date you requested ( " & CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX)) & " ) " & _
                             vbNewLine & vbNewLine & "Do wish to continue?", vbQuestion + vbYesNo + vbCritical, "Importing Data")
            If Committ = vbNo Then GoTo FixMe
        End If
        ReDim DataX(1 To lastRow, 1 To 19)
        Dim DateSerialx As String
        For i = 2 To lastRow
            .Application.Range("C" & i).Select      'Read Make into a temp String
            UnitTypeX = Trim(.Application.Selection.Text)
            .Application.Range("E" & i).Select      'Read VIN
            DataX(i, 1) = Trim(.Application.Selection.Text)
            .Application.Range("F" & i).Select      'Read Pool
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 2) = 0 Else DataX(i, 2) = .Application.Selection.Text
            .Application.Range("G" & i).Select      'Read AGMT
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 3) = 0 Else DataX(i, 3) = .Application.Selection.Text
            .Application.Range("H" & i).Select      'Read Miles Date
            If IsDate(.Application.Selection.Text) Then DataX(i, 4) = .Application.Selection.Text Else DataX(i, 4) = "1/1/1900"
            .Application.Range("I" & i).Select      'Read Miles
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 5) = 0 Else DataX(i, 5) = Nz(.Application.Selection.Text, 0)
            .Application.Range("J" & i).Select      'Read Miles Rate
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 6) = 0 Else DataX(i, 6) = .Application.Selection.Text
            .Application.Range("K" & i).Select      'Read Hours Date
            If IsDate(.Application.Selection.Text) Then DataX(i, 7) = .Application.Selection.Text Else DataX(i, 7) = "1/1/1900"
            .Application.Range("L" & i).Select      'Read Hours
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 8) = 0 Else DataX(i, 8) = .Application.Selection.Text
            .Application.Range("M" & i).Select      'Read Hours Rate
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 9) = 0 Else DataX(i, 9) = .Application.Selection.Text
            .Application.Range("N" & i).Select      'Read Stand By Date
            If IsDate(.Application.Selection.Text) Then DataX(i, 10) = .Application.Selection.Text Else DataX(i, 10) = "1/1/1900"
            'Read Stand By Rate
            .Application.Range("O" & i).Select
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 11) = 0 Else DataX(i, 11) = .Application.Selection.Text
            'Read Lift Gate Start Date
            .Application.Range("Q" & i).Select
            If IsDate(.Application.Selection.Text) Then DataX(i, 12) = .Application.Selection.Text Else DataX(i, 12) = "1/1/1900"
            'Read Lift Gate Cycles
            .Application.Range("R" & i).Select
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 13) = 0 Else DataX(i, 13) = .Application.Selection.Text
            'Read Lift Gate Rate
            .Application.Range("S" & i).Select
            If .Application.Selection.Text & "test" = "test" Then DataX(i, 14) = 0 Else DataX(i, 14) = .Application.Selection.Text
            ' Make and get UID - old version
            '.Application.Range("S" & i).Select
            '.Application.Selection.Formula = "=E" & i & "&G" & i
            ' DataX(i, 14) = .Application.Selection.Text

            ' Faster UID routine below
            ' Make and get UID
            .Application.Range("H" & i).Select

            DateSerialx = CLng(.Application.Selection.Value)
            .Application.Range("E" & i).Select
            DataX(i, 15) = .Application.Selection.Text & DateSerialx
            'Test UnitTypeX String to assign UnitType
            If UnitTypeX = "FRGH" Or UnitTypeX = "MACK" Or UnitTypeX = "NVSTR" Or UnitTypeX = "VOLV" Then DataX(i, 16) = "Tractor" Else DataX(i, 16) = "Trailer"
            DataX(i, 17) = ""
            DataX(i, 18) = Nz(DLookup("UnitInserv", "Units", "UnitVIN='" & DataX(i, 1) & "'"), "1/1/1900")
            Call MessageUser("Update", "VIN:" & DataX(i, 1) & "", "working on Unit " & i & " of " & lastRow - 1 & "")
        Next i

        Call MessageUser("Update", "Closing Excel", "...some house cleaning...")
        .Application.Selection.Columns.AutoFit
        'With .ActiveSheet
        .Range("A1:L" & lastRow).RemoveDuplicates


    End With    'end with XLApp
    xlApp.Application.DisplayAlerts = False
    xlApp.Application.Quit

    Call MessageUser("Update", "Posting data to ATLAAS Core", "...almost done...")
    DoCmd.SetWarnings False
    DoCmd.RunSQL ("Delete * From AIMDataTemp")
    Dim LogActX As String
    Dim ClientIDx, ClientGroupIDx As Long
    LogActX = CDate(DatePart("m", LogDateX) & "/1/" & DatePart("yyyy", LogDateX))
    SQLz = ""

    For i = 2 To lastRow
        SQLz = ""
        SQLz = SQLz & "INSERT INTO AIMDataTemp ("
        SQLz = SQLz & "UnitNumber, VIN, Pool, AGMT, ReadDateMiles, MilesLTD, MilesRate, ReadDateHours, HoursLTD, "
        SQLz = SQLz & "HoursRate, StandByDate, StandByRate, ReadDateLiftGate, LiftGateCyclesLTD, LiftGateRate, UID, "
        SQLz = SQLz & "UnitType, LeaseType, InServDate, LeaseTerm, LogAct ) "
        SQLz = SQLz & "SELECT '"
        SQLz = SQLz & Nz(DLookup("UnitUnitNum", "Units", "UnitVIN='" & DataX(i, 1) & "'"), "") & "', '"    'UnitNumber
        SQLz = SQLz & DataX(i, 1) & "', '"                      ' Vin
        SQLz = SQLz & DataX(i, 2) & "', '"                      ' Pool
        SQLz = SQLz & DataX(i, 3) & "', '"                      ' AGMT
        SQLz = SQLz & CDate(DataX(i, 4)) & "', "                ' Miles Date
        SQLz = SQLz & CDbl(DataX(i, 5)) & ", "                  ' Miles LTD
        SQLz = SQLz & CDbl(DataX(i, 6)) & ", '"                 ' Miles Rate
        SQLz = SQLz & CDate(DataX(i, 7)) & "', "                ' Hours Date
        SQLz = SQLz & CDbl(Nz(DataX(i, 8), 0)) & ", "           ' Hours LTD
        SQLz = SQLz & CDec(Nz(DataX(i, 9), 0)) & ", '"          ' Hours Rate
        SQLz = SQLz & CDate(DataX(i, 10)) & "', "               ' StandBy Date
        SQLz = SQLz & CDbl(Nz(DataX(i, 11), 0)) & ", '"         ' StandBy Rate
        SQLz = SQLz & CDate(DataX(i, 12)) & "', "               ' LiftGate Read Date
        SQLz = SQLz & CDbl(Nz(DataX(i, 13), 0)) & ", "          ' LiftGate Cycles LTD
        SQLz = SQLz & CDbl(Nz(DataX(i, 14), 0)) & ", '"         ' LiftGate Rate
        SQLz = SQLz & DataX(i, 15) & "', '"                     ' UID
        SQLz = SQLz & DataX(i, 16) & "', '"                     ' UnitType
        SQLz = SQLz & DataX(i, 17) & "', #"                     ' LeaseType
        SQLz = SQLz & CDate(Nz(DataX(i, 18), "1/1/1900")) & "#, "  ' InServDate
        SQLz = SQLz & "0, #"                                       ' LeaseTerm
        SQLz = SQLz & LogActX & "# "                               ' LogAct
        SQLz = SQLz & ";"

        DoCmd.RunSQL (SQLz)
    Next i

    i = 2
    Do While (True)
        ClientIDx = Nz(DMax("ClientID", "Units", "UnitVin='" & DataX(i, 1) & "'"), 0)
        ClientGroupIDx = Nz(DMax("ClientGroupID", "Units", "UnitVin='" & DataX(i, 1) & "'"), 0)
        If ClientIDx > 0 Then Exit Do
        i = i + 1
    Loop
    'Make Query to pull previous Month's MnR Data
    SQLz = "SELECT DISTINCT UnitID, VIN, LogActivityDate AS PrevLogActivityDate, MilesLTD AS PrevMilesLTD, HoursLTD AS PrevHoursLTD, LiftGateCyclesLTD AS PrevLiftGateLTD " & _
           "FROM MnRData " & _
           "WHERE (((MnRData.LogActivityDate)=#" & DateAdd("M", -1, LogActX) & "#));"
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "MnRDataDataLastRead"
    Set qdf = CurrentDb.CreateQueryDef("MnRDataDataLastRead", SQLz)
    On Error GoTo 0

    ' Delete exitsting data that is already there
    Call MessageUser("Update", "Clearing Exisiting Data", "...Only what is to be replaced...")

    SQLz = "DELETE * From MnRData " & _
           "WHERE EXISTS " & _
           "(Select * From AIMDataTemp Where AIMDataTemp.LogAct = MnRData.LogActivityDate AND AIMDataTemp.VIN = MnRData.VIN);"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Call MessageUser("Update", "Pushing To SQL Server", "...Just a bit more...")
    'Merge Tempdata with ATLAAS Core Fleet Info
    SQLz = ""
    SQLz = SQLz & "INSERT INTO MnRData ( UnitID, UnitNumber, VIN, POOL, AGMT, ReadDateMiles, PrevMilesLTD, Miles, MilesLTD, MilesRate, ReadDateHours, "
    SQLz = SQLz & "PrevHoursLTD, Hours, HoursLTD, HoursRate, StandByDate, StandByRate, ReadDateLiftGate, PrevLiftGateLTD, LiftGateCycles, LiftGateCyclesLTD, LiftGateRate, "
    SQLz = SQLz & "UID, LogActivityDate, UnitType, ClientID, ClientGroupID, LiftGateStartDate, ProviderID ) "
    SQLz = SQLz & "SELECT DISTINCT "
    SQLz = SQLz & "Nz([Units.UnitID],0) AS UnitID, "
    SQLz = SQLz & "AIMDataTemp.UnitNumber, "
    SQLz = SQLz & "AIMDataTemp.VIN, "
    SQLz = SQLz & "AIMDataTemp.POOL, "
    SQLz = SQLz & "AIMDataTemp.AGMT, "
    SQLz = SQLz & "AIMDataTemp.ReadDateMiles, "
    SQLz = SQLz & "Nz(MnRDataDataLastRead.PrevMilesLTD,0), "
    SQLz = SQLz & "Nz([MilesLTD]-[PrevMilesLTD],0) AS Miles, "
    SQLz = SQLz & "AIMDataTemp.MilesLTD, AIMDataTemp.MilesRate, AIMDataTemp.ReadDateHours, "
    SQLz = SQLz & "Nz(MnRDataDataLastRead.PrevHoursLTD,0), "
    SQLz = SQLz & "Nz(HoursLTD - [PrevHoursLTD],0) AS Hours, "
    SQLz = SQLz & "AIMDataTemp.HoursLTD, AIMDataTemp.HoursRate, AIMDataTemp.StandByDate, AIMDataTemp.StandByRate, AIMDataTemp.ReadDateLiftGate, "
    SQLz = SQLz & "Nz(MnRDataDataLastRead.PrevLiftGateLTD,0), "
    SQLz = SQLz & "Nz([LiftGateCyclesLTD]-[PrevLiftGateLTD],0) AS LiftGateCycles, "
    SQLz = SQLz & "AIMDataTemp.LiftGateCyclesLTD, AIMDataTemp.LiftGateRate, AIMDataTemp.UID, "
    SQLz = SQLz & "#" & LogActX & "# AS LogActivityDate, "
    SQLz = SQLz & "AIMDataTemp.UnitType, " & ClientIDx & ", " & ClientGroupIDx & ", AIMDataTemp.InServDate, 101 "
    SQLz = SQLz & "FROM (AIMDataTemp LEFT JOIN Units ON AIMDataTemp.VIN = Units.UnitVIN) "
    SQLz = SQLz & "LEFT JOIN MnRDataDataLastRead ON AIMDataTemp.VIN = MnRDataDataLastRead.VIN;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call PostTrailerMilesHours(LogActX, 101)
    Call TrailerMilesHoursDataValidator(LogActX, 101)
    Call SendImportEmailMsg("AIM", ClientGroupIDx, LogActX, TotalRows)
    Exit Function
FixMe:
    Call MessageUser("Close")
    xlApp.Application.DisplayAlerts = False
    xlApp.Application.Quit
    MsgBox "This file doesn't have the proper MNR data" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"
End Function
Function PostTrailerMilesHours(LogActX As String, ProviderIDx As Integer)
'Put Trailer Data into TrailerMilesHours table
    On erorr GoTo FixMe
    Dim SQLz   As String
    Call MessageUser("Close")
    Call MessageUser("Open", "Clearing Exisiting Trailer Data", "...Only what is to be replaced...")
    SQLz = "DELETE DISTINCTROW TrailerMilesHours.*, MnRData.ProviderID, MnRData.LogActivityDate " & _
           "FROM TrailerMilesHours INNER JOIN MnRData " & _
           "ON (TrailerMilesHours.LogActivityDate = MnRData.LogActivityDate) " & _
           "AND (TrailerMilesHours.UnitID = MnRData.UnitID) " & _
           "WHERE (((MnRData.ProviderID)=" & ProviderIDx & ") " & _
           "AND ((MnRData.LogActivityDate)=#" & LogActX & "#));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Call MessageUser("Update", "Posting Trailer Data", "...will be done in a moment...")
    SQLz = "INSERT INTO TrailerMilesHours ( UnitID, UnitNumber, VIN, LogActivityDate, ReadDate, Miles, CPM, MilesLTD, Hours, CPH, " & _
           "HoursLTD, StandByObservDate, StandByRate, LiftGateObservDate, LiftGateRate, ProviderID ) " & _
           "SELECT DISTINCTROW UnitID, UnitNumber, VIN, LogActivityDate, ReadDateMiles, Miles, Max(MilesRate) AS MaxMilesRate, MilesLTD, " & _
           "Hours, Max(HoursRate) AS MaxHoursRate, HoursLTD, StandByDate, Max(StandByRate) AS MaxStandByRate, LiftGateStartDate, LiftGateRate, ProviderID " & _
           "FROM MnRData " & _
           "WHERE LogActivityDate=#" & LogActX & "#  AND UnitType='Trailer' AND ProviderID=" & ProviderIDx & " " & _
           "GROUP BY UnitID, UnitNumber, VIN, LogActivityDate, ReadDateMiles, Miles, MilesLTD, Hours, HoursLTD, StandByDate, LiftGateStartDate, LiftGateRate, ProviderID;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Call MessageUser("Close")
    Exit Function
FixMe:
    Call MessageUser("Close")
    xlApp.Application.Quit
    MsgBox "Proc: Post Trailer Miles Hours Error" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"
End Function
Function TrailerMilesHoursDataValidator(LogActX As String, ProviderIDx As Integer)
    On erorr GoTo FixMe
    Dim SQLz   As String
    Call MessageUser("Close")
    Call MessageUser("Open", "Fix Missing UnitNums to Trailer Data", "...using Units Table...")
    SQLz = "UPDATE TrailerMilesHours INNER JOIN Units ON TrailerMilesHours.UnitID = Units.UnitID SET TrailerMilesHours.UnitNumber = [Units].[UnitUnitnum] " & _
           "WHERE (((TrailerMilesHours.UnitNumber) Is Null));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Call MessageUser("Update", "Validating Trailer Data", "...Setting Data Status...")
    SQLz = "UPDATE TrailerMilesHours INNER JOIN TrailerMilesHoursDataValidator ON " & _
           "(TrailerMilesHours.LogActivityDate = TrailerMilesHoursDataValidator.LogActivityDate) AND (TrailerMilesHours.UnitID = TrailerMilesHoursDataValidator.UnitID) " & _
           "SET TrailerMilesHours.OBCDataStatus = [TrailerMilesHoursDataValidator].[DataStatus], " & _
           "TrailerMilesHours.OBCDataDesc = [TrailerMilesHoursDataValidator].[DataStatusDesc] " & _
           "WHERE TrailerMilesHours.LogActivityDate = #" & LogActX & "# AND ProviderID=" & ProviderIDx & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True

    Call MessageUser("Close")
    Exit Function

FixMe:
    Call MessageUser("Close")
    xlApp.Application.Quit
    MsgBox "Proc: Data Validator Error" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"
End Function


Sub Delete_Recreate_TheWorkSheet()
    On Error GoTo ErrorHandler

    '=========================
    Dim strInFrontOfSheetName As String
    m_sNameOfOutPutWorkSheet_1 = "Dashboard_1"
    strInFrontOfSheetName = "CONTROL"    'create the new worksheet in front of this sheet

    '1] Clean up old data if it is still there
    GetRidOf_WorkSheet_IfItExists (m_sNameOfOutPutWorkSheet_1)

    CreateNewOutputWorkSheet m_sNameOfOutPutWorkSheet_1, strInFrontOfSheetName
    'Color the tab of the new worksheet
    ActiveWorkbook.Sheets(m_sNameOfOutPutWorkSheet_1).Tab.ColorIndex = 5

    'Select the worksheet that I started with
    Worksheets(strInFrontOfSheetName).Select

    '=========================
    Exit Sub

ErrorHandler:
    Select Case Err.Number
    Case Else
        MsgBox "One_Main - Error: " & Err.Number & " " & Err.Description
    End Select
End Sub

Sub GetRidOf_WorkSheet_IfItExists(sWorkSheetName_ForInitalData As String)
    On Error GoTo ErrorHandler

    '=========================

    If fn_WorkSheetExists(sWorkSheetName_ForInitalData) Then
        'Sheet Exists
         xlApp.Application.DisplayAlerts = False
         xlApp.Worksheets(sWorkSheetName_ForInitalData).Delete
         xlApp.Application.DisplayAlerts = True

    End If

    '=========================
    Exit Sub

ErrorHandler:
    Select Case Err.Number
    Case Else
        MsgBox "GetRidOf_WorkSheet_IfItExists - Error: " & Err.Number & " " & Err.Description
    End Select
End Sub


Function fn_WorkSheetExists(wsName As String) As Boolean
    On Error Resume Next
    fn_WorkSheetExists = Worksheets(wsName).Name = wsName
End Function


Sub CreateNewOutputWorkSheet(sWorkSheetName_ForOutputData As String, strInFrontOfSheetName As String)
    On Error GoTo ErrorHandler

    '=========================
    If fn_WorkSheetExists(sWorkSheetName_ForOutputData) Then
        'Sheet Exists
        xlApp.Application.DisplayAlerts = False
        xlApp.Worksheets(sWorkSheetName_ForOutputData).Delete
        xlApp.Application.DisplayAlerts = True
    End If

    Dim wsX    As Worksheet
    Set wsX = xlApp.Sheets.Add(Before:=Worksheets(strInFrontOfSheetName))

    wsX.Name = xlApp.sWorkSheetName_ForOutputData

    '=========================
    Exit Sub

ErrorHandler:
    Select Case Err.Number
    Case Else
        MsgBox "CreateNewOutputWorkSheet - Error: " & Err.Number & " " & Err.Description
    End Select
End Sub
Function OpenVendorDataFileDialog(InspectionNum As Integer)
Exit Function
'' -- disabled by Leo
    If Not VBA.Environ("HOMESHARE") Like "*FA-File*" Then
        MsgBox "Sorry, You're not on the Fleet Advantage Network", , "Not Local"
        Exit Function
    End If
    Dim f      As Object
    Set f = Application.FileDialog(3)
    f.InitialFileName = Environ("Userprofile") & "\Documents"
    f.AllowMultiSelect = False
    f.Filters.Clear
    f.Filters.Add "Excel Macro-Enabled", "*.xlsm"
    f.Filters.Add "Excel", "*.xlsx"
    f.Filters.Add "Excel 97", "*.xlx"
    f.Filters.Add "CSV", "*.csv"
    f.Show
    On Error GoTo FixMe
    ' MsgBox "file choosen = " & f.SelectedItems(1)
    Call GetVendorDataExcel(f.SelectedItems(1), InspectionNum)
    Exit Function
FixMe:
End Function
Public Function GetVendorDataExcel(sFile As String, ByVal InspectionNum As Integer)
    DoCmd.SetWarnings False
    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim SelectedRange, WorkingRange As Range
    Dim i, x   As Integer
    Dim TargetX, SQLz, DataX(), AssetX As String
    ' On Error GoTo 0
    Set xlApp = CreateObject("Excel.Application")
    'mFile = sFile
    If sFile Like "*.xlsm" Then sFile = sFile Else sFile = sFile & ".xlsm"
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .Visible = False
        Set xlWB = .Workbooks.Open(sFile, , False)
        DoCmd.SetWarnings True
        Dim lastRow As Long
        With .ActiveSheet
            lastRow = .Range("A" & .rows.Count).End(xlUp).Row
            If StartRow > 1 Then lastRow = lastRow + (StartRow - 1)
        End With
        ReDim DataX(1 To (lastRow - 8), 1 To 12)
        .Application.Range("A9").Select
        AssetX = .Application.Selection.Text
        TargetX = "Client:    " & DLookup("ClientShortName", "vw_SixKeys", "[AssetID]=" & AssetX & "")
        TargetX = TargetX & vbNewLine & "Schdule: " & DLookup("Schedule", "vw_SixKeys", "[AssetID]=" & AssetX & "")
        TargetX = TargetX & vbNewLine & "Group:    " & DLookup("UnitGroup", "vw_SixKeys", "[AssetID]=" & AssetX & "")
        TargetX = TargetX & vbNewLine & "Units In:   " & (lastRow - 8)
        If DLookup("FAID", "vw_SixKeys", "[AssetID]=" & AssetX & "") <> Forms!Open_Items.FAID Then TargetX = TargetX & vbNewLine & "Note: This Does NOT match the PO you have open!"
        Dim Committ, SetString As String
        Committ = MsgBox("You're Collecting Vendor Data For..." & vbNewLine & TargetX & vbNewLine & "Do wish to continue?", vbQuestion + vbYesNo + vbCritical, "Vendor Data")
        If Committ = vbNo Then
            Call MessageUser("Close")
            xlApp.Application.Quit
            Exit Function
        End If
        Call MessageUser("Open", "Getting Vendor Data", "...This will take a moment...", False)
        'AssetID
        For i = 9 To lastRow
            .Application.Range("A" & i).Select
            DataX(i - 8, 1) = .Application.Cells((i), 1).Text
            DataX(i - 8, 2) = .Application.Cells((i), 9).Text
            DataX(i - 8, 3) = .Application.Cells((i), 10).Text
            DataX(i - 8, 4) = .Application.Cells((i), 11).Text
            DataX(i - 8, 5) = .Application.Cells((i), 12).Text
            DataX(i - 8, 6) = .Application.Cells((i), 13).Text
            DataX(i - 8, 7) = .Application.Cells((i), 14).Text
            DataX(i - 8, 8) = .Application.Cells((i), 15).Text
            DataX(i - 8, 9) = .Application.Cells((i), 16).Text
            DataX(i - 8, 10) = .Application.Cells((i), 17).Text
            DataX(i - 8, 11) = .Application.Cells((i), 18).Text
            DataX(i - 8, 12) = .Application.Cells((i), 19).Text
            Call MessageUser("Update", "Getting AssetID: " & DataX(i - 8, 1) & "", "getting record: " & (i - 8) & " of:" & (lastRow - 8), False)
        Next i
    End With
    Call MessageUser("Update", "Posting to tables", "...Almost done...")
    xlApp.Application.Quit
    DoCmd.SetWarnings False
    DoCmd.DeleteObject acTable, "OpenVendorDataTemp"
    SQLz = "CREATE TABLE OpenVendorDataTemp ( " & _
           "AssetID INTEGER CONSTRAINT PK_tblCustomers PRIMARY KEY, " & _
           "[UnitVIN] TEXT(50), " & _
           "[UnitBuildDate] DATETIME, " & _
           "[UnitEstDelDate] DATETIME, " & _
           "[UnitActualDelDate] DATETIME, " & _
           "[UnitEngineSerial] TEXT(50), " & _
           "[UnitTransSerial] TEXT(50), " & _
           "[UnitFASerial] TEXT(50), " & _
           "[UnitRAFSerial] TEXT(50), " & _
           "[UnitRARSerial] TEXT(50), " & _
           "[Ref_Serial] TEXT(50), " & _
           "[Ref_Engine] TEXT(50))"
    DoCmd.RunSQL (SQLz)
    For i = 1 To (lastRow - 8)
        SQLz = "INSERT INTO OpenVendorDataTemp (AssetID, UnitVIN, UnitBuildDate, UnitEstDelDate, UnitActualDelDate, UnitEngineSerial, UnitTransSerial, UnitFASerial, UnitRAFSerial, UnitRARSerial, Ref_Serial,Ref_Engine) " & _
               "SELECT " & DataX(i, 1) & ", '" & DataX(i, 2) & "', " & _
               "'" & DataX(i, 3) & "', '" & DataX(i, 4) & "', '" & DataX(i, 5) & "', " & _
               "'" & DataX(i, 6) & "', '" & DataX(i, 7) & "', " & _
               "'" & DataX(i, 8) & "', '" & DataX(i, 9) & "', '" & DataX(i, 10) & "', " & _
               "'" & DataX(i, 11) & "', '" & DataX(i, 12) & "' ;"
        DoCmd.RunSQL (SQLz)
    Next i
    DoCmd.OpenQuery "OpenVendorTrackingUpdate_AllDates"
    SQLz = "UPDATE OpenVendorDataTemp INNER JOIN Units ON OpenVendorDataTemp.AssetID = Units.UnitID " & _
           "SET Units.UnitVIN = [OpenVendorDataTemp].[UnitVIN], Units.UnitBuildDate = [OpenVendorDataTemp].[UnitBuildDate], " & _
           "Units.UnitEstDelDate = [OpenVendorDataTemp].[UnitEstDelDate], Units.UnitActualDelDate = [OpenVendorDataTemp].[UnitActualDelDate], " & _
           "Units.UnitEngineSerial = [OpenVendorDataTemp].[UnitEngineSerial], Units.UnitTransSerial = [OpenVendorDataTemp].[UnitTransSerial], " & _
           "Units.UnitFASerial = [OpenVendorDataTemp].[UnitFASerial], Units.UnitRAFSerial = [OpenVendorDataTemp].[UnitRAFSerial], Units.UnitRARSerial = [OpenVendorDataTemp].[UnitRARSerial];"
    DoCmd.RunSQL (SQLz)
    DoCmd.SetWarnings True
    Forms!Open_Items.Open_Items_SubForm.Requery
    Forms!Open_Items.Open_Items_SubForm2.Requery
    Forms!Open_Items.Submission_List.Requery
    Call MessageUser("Close")
    Exit Function
FixMe:
    xlApp.Application.Quit
    MsgBox "This file doesn't have the proper inspection data" & vbNewLine & "Erorr Code: " & Err, vbCritical, "ATLAAS Info"
    Call MessageUser("Close")
End Function
Public Function SendImportEmailMsg(ProviderType As String, ClientGroupIDx As Long, LogDate As String, DataRows As Long)
    Dim msg1   As String
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'>'" & ProviderType & "' data was imported</p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Fleet Services / IT Dept.<br /><br />"
    'Msg1 = Msg1 & "All of the units in Group (" & GrpX & ") have been marked Status 'A':<br /><br />"
    msg1 = msg1 & "Client Group: <B>" & DLookup("ClientGroupName", "ClientGroups", "ClientGroupID=" & ClientGroupIDx) & "</B><br />"
    msg1 = msg1 & "Log Activity Date: <B>" & LogDate & "</B><br />"
    msg1 = msg1 & "Provider: <B>" & ProviderType & "</B><br />"
    msg1 = msg1 & "Data Rows: <B>" & Format(DataRows, "#,###") & "</B><br />"
    msg1 = msg1 & "Date Processed: <B>" & Now() & "</B><br />"
    msg1 = msg1 & "User who ran importer: <B>" & FAUserFullName & "</B><br /><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"
    Call SendEmailMessage("cvanatta@FleetAdvantage.com;mhendrix@FleetAdvantage.com;gbukowski@FleetAdvantage.com", "ATLAAS Core Message: " & ProvideerType & " data was imported", msg1, False, True, False, "IT@fleetadvantage.com")
    Exit Function
End Function
