Attribute VB_Name = "Export2Excel"
Option Compare Database
Public VendorAdr(4), VendorContact(4), DecalVendor(12) As String
Public RST As DAO.Recordset
Public cnt As Integer
Public SQLz As String
'Place the "Public vStatusBar As Variant" at the top of a public module, not a form module.
Public vStatusBar As Variant

Public Function FormatXL(ByVal TableNamex As String, sFile As String, ByVal sTabName As String, Optional ByVal sClose As String, _
                         Optional ByVal sUnitType As String, Optional ByVal sUnitCount As Integer, Optional ByVal sVendorID As Integer, _
                         Optional ByVal SumLine As Boolean, Optional PrintOpts As String)
    DoCmd.SetWarnings False
    Dim xlApp As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim mFile As String
    If sUnitCount = 0 Then sUnitCount = 300 Else sUnitCount = sUnitCount + 2

    Set xlApp = CreateObject("Excel.Application")
    mFile = sFile
    If sFile Like "*.xlsx" Then sFile = sFile Else sFile = sFile & ".xlsx"
    Set xlApp = CreateObject("Excel.Application")
    If FileExists(sFile) Then
        Kill sFile
    End If
    If FileExists(mFile & ".xlsm") Then
        Kill (mFile & ".xlsm")
    End If
    DoCmd.TransferSpreadsheet acExport, , TableNamex, sFile
    DoCmd.SetWarnings True


    Application.SetOption "Show Status Bar", True
    vStatusBar = SysCmd(acSysCmdSetStatus, "Formatting export file... please wait.")
    Set xlSheet = xlApp.Workbooks.Open(sFile).Sheets(1)
    With xlApp
        .Application.Visible = True
        .Application.Sheets(TableNamex).Select
        .Application.ActiveSheet.Name = sTabName
        .Application.Cells.Select
        ' .Application.Selection.ClearFormats
        .Application.Selection.Font.Name = "Calibri"
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
        .Application.ActiveWorkbook.Save

CustomFormat:
        'Set row colors
        For i = 1 To (sUnitCount - 1) Step 2
            .Application.Range("A" & i & ":AH" & i).Select
            .Application.Selection.Interior.Color = RGB(240, 240, 240)    ' light grey
        Next i
        i = 0
        If Not IsNull(sUnitType) Then


            Dim FomulaX, StringX As String

            'Set the Columns:
            .Application.Range("A1" & ":A" & ((sUnitCount - 1))).Select
            .Application.Selection.Interior.ColorIndex = 15
            .Application.Selection.Font.Bold = False
            .Application.Range("B1" & ":B" & ((sUnitCount - 1))).Select
            .Application.Selection.HorizontalAlignment = xlCenter
            .Application.Range("C1" & ":C" & ((sUnitCount - 1))).Select
            .Application.Selection.Interior.ColorIndex = 15
            .Application.Selection.HorizontalAlignment = xlCenter
            .Application.Selection.Font.Bold = False
            .Application.Range("D1" & ":D" & ((sUnitCount - 1))).Select
            .Application.Selection.HorizontalAlignment = xlCenter
            .Application.Range("H1" & ":H" & ((sUnitCount - 1))).Select
            .Application.Selection.HorizontalAlignment = xlCenter
            .Application.Range("T1" & ":T" & ((sUnitCount - 1))).Select
            .Application.Selection.Interior.ColorIndex = 24
            .Application.Selection.Font.Bold = True
            .Application.Selection.Style = "Currency"
            .Application.Range("I1" & ":U" & ((sUnitCount - 1))).Select
            .Application.Selection.Font.Bold = False
            .Application.Selection.Font.Size = 10
            .Application.Selection.Borders.Weight = xlThin

            'Make Header:
            .Application.rows("1:1").Select
            .Application.Selection.Interior.ColorIndex = 25
            .Application.Selection.Font.Color = vbWhite
            'Set Sum/Count Cells
            .Application.Range("A1").Select
            .Application.Selection.EntireRow.Insert
            .Application.Range("A1").Select
            .Application.Range("A1").formula = "=SUBTOTAL(103,A3:A" & sUnitCount & ")"
            .Application.Selection.Interior.ColorIndex = 15
            .Application.Selection.Borders.Weight = xlThin
            .Application.Selection.Interior.ColorIndex = 15
            .Application.Selection.Borders.Weight = xlThin

            .Application.Range("T1").Select
            .Application.Range("T1").formula = "=Sum(T3:T" & sUnitCount & ")"
            .Application.Selection.Interior.ColorIndex = 24
            .Application.Selection.Borders.Weight = xlThin

            StringX = """Needed """
            If SumLine = True Then
                .Application.Range("I1").formula = "=" & StringX & " & A1 - COUNTA(I3:I" & sUnitCount & ")"
                .Application.Range("M1").formula = "=" & StringX & " & A1 - COUNTA(M3:M" & sUnitCount & ")"
                .Application.Range("N1").formula = "=" & StringX & " & A1 - COUNTA(N3:N" & sUnitCount & ")"
                .Application.Range("O1").formula = "=" & StringX & " & A1 - COUNTA(O3:O" & sUnitCount & ")"
                .Application.Range("P1").formula = "=" & StringX & " & A1 - COUNTA(P3:P" & sUnitCount & ")"
                .Application.Range("Q1").formula = "=" & StringX & " & A1 - COUNTA(Q3:Q" & sUnitCount & ")"
                .Application.Range("R1").formula = "=" & StringX & " & A1 - COUNTA(R3:R" & sUnitCount & ")"
                .Application.Range("S1").formula = "=" & StringX & " & A1 - COUNTA(S3:S" & sUnitCount & ")"
                .Application.Range("I1:I1").Select
                .Application.Selection.Borders.Weight = xlThin
                .Application.Range("M1:S1").Select
                .Application.Selection.Borders.Weight = xlThin
            End If
            .Application.Columns().Select
            .Application.Selection.Columns.AutoFit
            .Application.Range("M1:T1").Select
            .Application.Selection.ColumnWidth = 18
            .Application.Range("V1:V1").Select
            .Application.Selection.ColumnWidth = 28

            'Set non-Lock Columns
            .Application.Columns("I:S").Locked = False
            .Application.Columns("V").Locked = False
            .Application.Columns("AE:AH").Locked = False

            'Hide the Columns:
            .Application.Range("Z1:AH1").Select
            .Application.Selection.EntireColumn.Hidden = True

            If sUnitType Like "*Trailer*" Then
                .Application.Range("M1:O1").Select
                .Application.Selection.EntireColumn.Hidden = True
            End If
            If sUnitType Like "*Refrigerated*" Then
                '       .Application.Range("J1:Q1").Select
                '       .Application.Selection.ColumnWidth = 18
            Else
                .Application.Range("R1:R1").Select
                .Application.Selection.EntireColumn.Hidden = True

            End If
            .Application.Range("X1:Y1").Select
            .Application.Selection.EntireColumn.Hidden = True
            .Application.Range("U1:U1").Select
            .Application.Selection.EntireColumn.Hidden = True

            'InsertRows for header:
            .Application.Range("A1").Select
            For i = 1 To 5
                .Application.Selection.EntireRow.Insert
            Next i
            .Application.rows((i - 1) & ":1").Select
            .Application.Selection.Interior.ColorIndex = xlNone
            .Application.Selection.Borders.LineStyle = xlNone

            'SetHeaderText:
            .Application.Range("A1") = "Vendor: "
            .Application.Range("A2") = "Address: "
            .Application.Range("A3") = "City, St Zip: "
            .Application.Range("A4") = "Phone: "
            .Application.Range("A1:A4").Select
            .Application.Selection.Font.Name = "Calibri"
            .Application.Selection.HorizontalAlignment = xlRight
            .Application.Selection.Font.Bold = False
            .Application.Range("E1") = "Name: "
            .Application.Range("E2") = "Email: "
            .Application.Range("E3") = "Phone: "
            .Application.Range("E4") = "Mobile: "
            .Application.Range("E1:E4").Select
            .Application.Selection.Font.Name = "Calibri"
            .Application.Selection.HorizontalAlignment = xlRight
            .Application.Selection.Font.Bold = False
            .Application.Selection.Font.Name = "Calibri"
            .Application.Selection.HorizontalAlignment = xlRight


            'Get Vendor
            Call GetVendor(Forms!Open_Items!VendorID)
            For i = 1 To 4
                .Application.Range("B" & i & ":D" & i).Select
                .Application.Selection.Merge
                .Application.Range("B" & i) = VendorAdr((i - 1))
                If i > 3 Then
                    .Application.Selection.NumberFormat = "(###) ### - ####"  'Set PhoneNum Format
                    .Application.Selection.HorizontalAlignment = xlLeft
                End If
            Next i

            Call GetVendorContactInfo(Forms!Open_Items!VContactID)
            For i = 1 To 4
                .Application.Range("F" & i & ":H" & i).Select
                .Application.Selection.Merge
                .Application.Range("F" & i) = VendorContact((i - 1))
                .Application.Selection.Font.Bold = False
                If i > 2 Then
                    .Application.Selection.NumberFormat = "(###) ### - ####"  'Set PhoneNum Format
                    .Application.Selection.HorizontalAlignment = xlLeft
                End If
            Next i
            'Get Decal Vendor
            .Application.Range("V2") = "N/A"
            .Application.Range("V1:AI4").Select
            .Application.Selection.Font.Name = "Calibri"
            Call GetDecalVendor(Forms!Open_Items!FAID)
            For i = 1 To 4
                .Application.Range("V" & i & ":V" & i).Select
                .Application.Selection.Merge
                .Application.Range("V" & i) = DecalVendor((i))
                If i > 3 Then
                    .Application.Selection.NumberFormat = "(###) ### - ####"  'Set PhoneNum Format
                    .Application.Selection.HorizontalAlignment = xlLeft
                End If
            Next i
            'Get 2nd Decal Vendor
            If cnt > 1 Then
                For i = 1 To 4
                    .Application.Range("W" & i & ":W" & i).Select
                    .Application.Selection.Merge
                    .Application.Range("W" & i) = DecalVendor((i + 4))
                    If i > 3 Then
                        .Application.Selection.NumberFormat = "(###) ### - ####"  'Set PhoneNum Format
                        .Application.Selection.HorizontalAlignment = xlLeft
                    End If
                Next i
            End If
            'Get 3rd Decal Vendor
            If cnt > 2 Then
                For i = 1 To 4
                    .Application.Range("AI" & i & ":AI" & i).Select
                    .Application.Selection.Font.Name = "Calibri"
                    .Application.Selection.Merge
                    .Application.Range("AI" & i) = DecalVendor((i + 8))
                    If i > 3 Then
                        .Application.Selection.NumberFormat = "(###) ### - ####"  'Set PhoneNum Format
                        .Application.Selection.HorizontalAlignment = xlLeft
                    End If
                Next i
            End If


            'Insert Client Name
            .Application.Range("A1").Select
            .Application.Selection.EntireRow.Insert
            .Application.rows("1:1").Select
            .Application.Selection.Interior.ColorIndex = xlNone
            .Application.Range("A1:L1").Select
            .Application.Selection.Merge
            .Application.Selection.VerticalAlignment = xlCenter
            .Application.Selection.Font.Bold = True
            .Application.Selection.Font.Size = 20
            .Application.Selection.RowHeight = 20
            .Application.Range("A1") = Forms!Open_Items!Client & " | PO " & Forms!Open_Items!PO & " | FAID " & Forms!Open_Items!FAID
            '.Application.Selection.EntireRow.Insert
            .Application.Range("V1") = "Decal Vendor(s):"

            Dim vbProj As VBIDE.VBProject
            Dim VBComp As VBIDE.VBComponent
            Dim CodeMod As VBIDE.CodeModule

            Dim chkRef As Reference
            Dim BoolExists As Boolean
            Dim CodeX As String
            Dim LineNum As Long
            Dim DQUOTE As String
            DQUOTE = """"
            'Const DQUOTE = """"    ' one " character
            Set vbProj = .Application.ActiveWorkbook.VBProject

            vbProj.References.AddFromFile "C:\WINDOWS\system32\msxml3.dll"

            Set VBComp = vbProj.VBComponents("ThisWorkbook")
            Set CodeMod = VBComp.CodeModule

            '             For Each chkRef In VBProj.References
            '                 If chkRef.Name = "VBScript_RegExp_55" Then BoolExists = True
            '             Next chkRef
            '      If BoolExists <> True Then

            'Define Code to be inserted
            CodeX = ""
            'CodeX = CodeX & "Private Sub Workbook_Open()" & vbCrLf
            'CodeX = CodeX & "If Sheet1.TransmitData.Caption Like " & DQUOTE & "*Macro*" & DQUOTE & "then" & vbCrLf
            'CodeX = CodeX & "  Sheet1.TransmitData.Caption = " & DQUOTE & "- Click HERE -" & DQUOTE & "& vbNewLine & " & DQUOTE & "to Send Data To " & DQUOTE & "& vbNewLine & " & DQUOTE & "Fleet Advantage" & DQUOTE & vbCrLf
            'CodeX = CodeX & "End If" & vbCrLf
            'CodeX = CodeX & "End Sub" & vbCrLf
            'With CodeMod
                'LineNum = .CreateEventProc("Open", "ThisWorkBook")
            '    LineNum = LineNum + 1
            '    .InsertLines LineNum, CodeX
            'End With

            Set VBComp = vbProj.VBComponents("Sheet1")
            Set CodeMod = VBComp.CodeModule
            CodeX = ""
            CodeX = CodeX & "Dim OldVal, Response As String" & vbCrLf
            CodeX = CodeX & "Public Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf
            CodeX = CodeX & "If Target.Column >= 10 And Target.Column <= 12 Then" & vbCrLf
            CodeX = CodeX & "  IF Not isnull(Target.Text) then OldVal = Target.Text" & vbCrLf
            CodeX = CodeX & "End If" & vbCrLf
            CodeX = CodeX & "End Sub" & vbCrLf
            LineNum = 3
            With CodeMod
                'LineNum = .CreateEventProc("Change", "Worksheet")
                LineNum = LineNum + 1
                .InsertLines LineNum, CodeX
            End With


            CodeX = ""
            CodeX = CodeX & "If Target.Row > 7 Then" & vbCrLf
            'CodeX = CodeX & "Public Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf
            CodeX = CodeX & " If Target.Column >= 10 And Target.Column <= 12 Then " & vbCrLf
            CodeX = CodeX & "  Cells(Target.Row, Target.Column).Interior.ColorIndex = 36" & vbCrLf
            CodeX = CodeX & "    If IsDate(Target.Text) Or Target.Text & " & DQUOTE & "test" & DQUOTE & " = " & DQUOTE & "test" & DQUOTE & " Then" & vbCrLf
            CodeX = CodeX & "      If Target.Text & " & DQUOTE & "test" & DQUOTE & " = " & DQUOTE & "test" & DQUOTE & " Then Exit Sub" & vbCrLf
            CodeX = CodeX & "      If Target.Text = OldVal Then Exit Sub" & vbCrLf
            CodeX = CodeX & "      If Target.Column = 10 Then Cells(Target.Row," & DQUOTE & "AE" & DQUOTE & ") = OldVal" & vbCrLf
            CodeX = CodeX & "      If Target.Column = 11 Then Cells(Target.Row," & DQUOTE & "AF" & DQUOTE & ") = OldVal" & vbCrLf
            CodeX = CodeX & "      If Target.Column = 12 Then" & vbCrLf
            CodeX = CodeX & "          if OldVal =" & DQUOTE & DQUOTE & " then Cells(Target.Row," & DQUOTE & "AG" & DQUOTE & ") =" & DQUOTE & "1/1/1900" & DQUOTE & " "
            CodeX = CodeX & "ELSE Cells(Target.Row," & DQUOTE & "AG" & DQUOTE & ") = OldVal" & vbCrLf
            CodeX = CodeX & "      End If " & vbCrLf
            CodeX = CodeX & "       Exit Sub" & vbCrLf
            CodeX = CodeX & "    Else" & vbCrLf
            CodeX = CodeX & "       MsgBox " & DQUOTE & "Sorry, but '" & DQUOTE & " & Target.Text & " & DQUOTE & "' is not a DATE. | Try again" & DQUOTE & "" & vbCrLf
            CodeX = CodeX & "       If IsNull(OldVal) Then Target = " & DQUOTE & DQUOTE & " Else Target = OldVal" & vbCrLf
            CodeX = CodeX & "          Cells(Target.Row, Target.Column).Interior.ColorIndex = 0" & vbCrLf
            CodeX = CodeX & "          Cells(Target.Row, Target.Column).Select " & vbCrLf
            CodeX = CodeX & "    End If " & vbCrLf
            CodeX = CodeX & " End If" & vbCrLf
            CodeX = CodeX & "End If" & vbCrLf
            CodeX = CodeX & "If Target.Row > 7 And Target.Column >= 13 And Target.Column <= 19 Then Cells(Target.Row, Target.Column).Interior.ColorIndex = 35" & vbCrLf
            LineNum = 3
            With CodeMod
                LineNum = .CreateEventProc("Change", "Worksheet")
                LineNum = LineNum + 1
                .InsertLines LineNum, CodeX
            End With
            'Insert SendButton Button
         '   Dim SendButton As Object

          '  Set SendButton = .Application.ActiveSheet.OLEObjects.Add _
                             (ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=740, Top:=19, Width:=80, Height:=60)
          '  .Application.ActiveSheet.OLEObjects(1).Object.Font.Size = 10
          '  .Application.ActiveSheet.OLEObjects(1).Object.caption = "Enable Macros" & vbCrLf & " To Send Data to" & vbCrLf & "Fleet Advantage"
          '  With SendButton
          '      .Name = "TransmitData"
                '          .Top = .Application.Selection
                '          .Left = .Application.Selection
                '          .Width = .Application.Selection
                '          .Height = .Application.Selection
                '.Visible = False
           ' End With
            'Define Code to be inserted
           ' CodeX = ""
           ' CodeX = CodeX & " If NOW() > #" & Date + 150 & "# Then " & vbCrLf
           ' CodeX = CodeX & "    MsgBox " & DQUOTE & "Tracking sheet Expired, Contact Fleet Advantage" & DQUOTE & vbCrLf & "  Exit Sub" & vbCrLf
           ' CodeX = CodeX & " End If" & vbCrLf
           ' CodeX = CodeX & " Response = MsgBox(" & DQUOTE & "Please Confrim DATA Transmission to Fleet Advantage?" & DQUOTE & ", vbQuestion + vbYesNo, " & DQUOTE & "Send Data?" & DQUOTE & ")" & vbCrLf
           ' CodeX = CodeX & "   If Response <> vbYes Then" & vbCrLf
           ' CodeX = CodeX & "      Exit Sub" & vbCrLf
           ' CodeX = CodeX & "   End If" & vbCrLf
           ' CodeX = CodeX & "Dim X As Long, G As Long, FF As Long, S() As String, DataStream As String, CleanMe As String" & vbCrLf
           ' CodeX = CodeX & "Dim PostX, RequestX As String" & vbCrLf
           ' CodeX = CodeX & "Dim FAResult As New MSXML2.DOMDocument" & vbCrLf
           ' CodeX = CodeX & "Dim FAService As New MSXML2.XMLHTTP" & vbCrLf
           ' CodeX = CodeX & "Range(" & DQUOTE & "A8:AH" & (sUnitCount + 6) & DQUOTE & ").Select" & vbCrLf
           ' CodeX = CodeX & "ReDim S(1 To Selection.Rows.Count)" & vbCrLf
           ' CodeX = CodeX & "DataStream = " & DQUOTE & "VendorID=" & sVendorID & "&Payload=" & DQUOTE & "" & vbCrLf
           ' CodeX = CodeX & "  For X = 1 To (Selection.Rows.Count)" & vbCrLf
           ' CodeX = CodeX & "      CleanMe = Join(WorksheetFunction.Transpose(WorksheetFunction. _ " & vbCrLf
           ' CodeX = CodeX & "                Transpose(Selection.Rows(X).Value)), " & DQUOTE & "|" & DQUOTE & ")" & vbCrLf
           ' CodeX = CodeX & "      CleanMe = Replace(CleanMe, " & DQUOTE & "," & DQUOTE & ", " & DQUOTE & DQUOTE & ")" & vbCrLf
           ' CodeX = CodeX & "      CleanMe = Replace(CleanMe, " & DQUOTE & "$" & DQUOTE & ", " & DQUOTE & DQUOTE & ")" & vbCrLf
           ' CodeX = CodeX & "      CleanMe = Replace(CleanMe, " & DQUOTE & DQUOTE & DQUOTE & DQUOTE & ", " & DQUOTE & "'" & DQUOTE & ")" & vbCrLf
           ' CodeX = CodeX & "      CleanMe = Replace(CleanMe, " & DQUOTE & "|" & DQUOTE & ", " & DQUOTE & "," & DQUOTE & ")" & vbCrLf
            'CodeX = CodeX & "      S(X) = CleanMe" & vbCrLf
            'CodeX = CodeX & "      DataStream = DataStream & S(X) & " & DQUOTE & "|" & DQUOTE & vbCrLf
            'CodeX = CodeX & "  Next" & vbCrLf
            ''CodeX = CodeX & "      PostX = " & DQUOTE & "http://services.fleetadvantage.net/api/Vendor/PostCsvData" & DQUOTE & "" & vbCrLf
            ''CodeX = CodeX & "   RequestX = " & DQUOTE & "http://services.fleetadvantage.net/api/Vendor/" & sVendorID & "/lastupdate" & DQUOTE & "" & vbCrLf
            ' +++++++++++++++++++++++++++ The URL's below were changed (from those above) 8/11/2015 for Azure
            'CodeX = CodeX & "      PostX = " & DQUOTE & "http://services.fleetadvantage.net/api/Vendor.asmx/PostCsvData " & DQUOTE & "" & vbCrLf
            'CodeX = CodeX & "   RequestX = " & DQUOTE & "http://services.fleetadvantage.net/api/Vendor.asmx/GetLastUpdateDate" & DQUOTE & "" & vbCrLf
            'CodeX = CodeX & "   FAService.Open " & DQUOTE & "POST" & DQUOTE & ", PostX, False" & vbCrLf
            'CodeX = CodeX & "   FAService.setRequestHeader " & DQUOTE & "Accept" & DQUOTE & ", " & DQUOTE & "text/html,application/xhtml+xml,*/*;q=0.8" & DQUOTE & "" & vbCrLf
            'CodeX = CodeX & "   FAService.setRequestHeader " & DQUOTE & "Content-Type" & DQUOTE & ", " & DQUOTE & "application/x-www-form-urlencoded" & DQUOTE & "" & vbCrLf
            'CodeX = CodeX & "   FAService.Send (DataStream)" & vbCrLf
            'CodeX = CodeX & "If FAService.responseText Like " & DQUOTE & "*true*" & DQUOTE & " Then" & vbCrLf
            'CodeX = CodeX & "    MsgBox " & DQUOTE & "Data Sent to Fleet Advantage" & DQUOTE & ", ," & DQUOTE & "Confrimed Data Sent" & DQUOTE & vbCrLf
            'CodeX = CodeX & " Me.TransmitData.Caption = " & DQUOTE & "Send Data To " & DQUOTE & "& vbNewLine & " & DQUOTE & "Fleet Advantage" & DQUOTE & " & vbNewLine & " & DQUOTE & "-last sent-" & DQUOTE & "& vbNewLine & Date" & vbCrLf
            'CodeX = CodeX & "  Range(" & DQUOTE & "AE9:AG" & (sUnitCount + 8) & DQUOTE & ").ClearContents" & vbCrLf
            'CodeX = CodeX & "Else" & vbCrLf
            'CodeX = CodeX & "    MsgBox " & DQUOTE & "Data Error, Data NOT sent; Please contact Fleet Advantage." & DQUOTE & vbCrLf
            'CodeX = CodeX & "End If" & vbCrLf
            'CodeX = CodeX & "  Range(" & DQUOTE & "A1" & DQUOTE & ").Select " & vbCrLf

            'With CodeMod
            '    LineNum = .CreateEventProc("Click", "TransmitData")
            '    LineNum = LineNum + 1
            '    .InsertLines LineNum, CodeX
            'End With
           ' .Application.ActiveSheet.OLEObjects(1).Object.caption = "Enable Macros" & vbCrLf & " To Send Data to" & vbCrLf & "Fleet Advantage"
            ''.Application.ActiveSheet.OLEObjects(1).Object.Caption = "- Click Here -" & vbCrLf & "to Send Data to" & vbCrLf & "Fleet Advantage"

InsertNote:
            'Insert Text box for Notes Area
             .Application.VBE.MainWindow.Visible = False
            Dim shp As Shape
            Set shp = .Application.ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 420, 20, 200, 62)    ' add shape
            With shp
                .Name = "Notesbox"
                ' .TextFrame.Selection.Characters(start:=1, Length:=216).Font.Name = "Calibri"
                .TextFrame.Characters.Text = "Notes:"    ' add text to display
                ' .TextFrame.AutoSize = True ' turn on autosize
                '.Fill.ForeColor.RGB = RGB(255, 255, 204)    'choose fill color
               ' .Line.Weight = 1    ' adjust width
               ' .Line.ForeColor.RGB = RGB(255, 0, 18)    ' choose color
               ' .Line.DashStyle = msoLineSolid    ' choose style
            End With

            .Application.Sheets(sTabName).Select
            .Application.ActiveSheet.Protect Password = "Fleet", _
                                             DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, _
                                             AllowFiltering:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
            .Application.Range("A1").Select
            .Application.ActiveWorkbook.Save
            GoTo CloseMe
NotCustom:
            .Application.Columns().Select
            .Application.Selection.Columns.AutoFit
            .Application.Range("A1").Select


        End If
CloseMe:
        .Application.ActiveWorkbook.Save
        .ActiveWorkbook.SaveAs fileName:=mFile & ".xlsm", FileFormat:=52
        Kill sFile
        If sClose = "Close" Then
            .Application.ActiveWorkbook.close
            .Quit
        End If
    End With


Cleanup:

    Set xlApp = Nothing
    Set xlSheet = Nothing
    vStatusBar = SysCmd(acSysCmdClearStatus)
    Exit Function
Err_FormatXL:
    If Err.Number = 70 Then
        MsgBox Err.Description & ": The Excel File (" & sFile & ") is already open. Please Close it or Save it as Different Filename", , "FAAS System"
    Else
        MsgBox Err.Number & " - " & Err.Description, , "FAAS System"
    End If
    GoTo Cleanup
End Function

Public Sub GetVendor(VendorIDx As Integer)
    SQLz = "Select * From Vendors WHERE VendorID =" & VendorIDx & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    RST.MoveFirst
    VendorAdr(0) = RST![VendorName]
    VendorAdr(1) = RST![VendorAddress]
    VendorAdr(2) = RST![VendorCity] & " " & RST![VendorState] & " " & RST![VendorZip]
    VendorAdr(3) = Nz(RST![VendorPhone], "-")
    RST.close
    Set RST = Nothing
End Sub
Public Sub GetVendorContactInfo(ContactIDx As Integer)
    SQLz = "SELECT * FROM VendorContacts WHERE ContactID = " & ContactIDx & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then
        VendorContact(0) = "N/A"
        VendorContact(1) = ""
        VendorContact(2) = ""
        VendorContact(3) = ""
        GoTo Cleanup
    End If
    RST.MoveLast
    RST.MoveFirst
    VendorContact(0) = RST!contactFirst & " " & RST!contactLast & ", " & RST!contactTitle
    VendorContact(1) = Nz(RST!contactEmail, "-")
    VendorContact(2) = Nz(RST!contactPrimaryPhone, "-")
    VendorContact(3) = Nz(RST!contactCell, "-")
Cleanup:
    RST.close
    Set RST = Nothing
End Sub

Public Sub GetDecalVendor(FAIDx As String)
    SQLz = "SELECT * FROM Open_Items_DecalVendors WHERE FAID = '" & FAIDx & "';"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    If RST.RecordCount < 1 Then
        DecalVendor(1) = "N/A"
        DecalVendor(2) = ""
        DecalVendor(3) = ""
        DecalVendor(4) = ""
        cnt = 0
        GoTo Cleanup
    End If

    RST.MoveLast
    RST.MoveFirst

    cnt = RST.RecordCount
    DecalVendor(1) = RST!VendorName
    DecalVendor(2) = RST!contactFirst & " " & RST!contactLast & ", " & RST!contactTitle
    DecalVendor(3) = Nz(RST!contactEmail, "-")
    DecalVendor(4) = Nz(RST!contactPrimaryPhone, "-")
    If cnt > 1 Then
        RST.MoveNext
        DecalVendor(5) = RST!VendorName
        DecalVendor(6) = RST!contactFirst & " " & RST!contactLast & ", " & RST!contactTitle
        DecalVendor(7) = Nz(RST!contactEmail, "-")
        DecalVendor(8) = Nz(RST!contactPrimaryPhone, "-")
        If cnt > 2 Then
            RST.MoveNext
            DecalVendor(9) = RST!VendorName
            DecalVendor(10) = RST!contactFirst & " " & RST!contactLast & ", " & RST!contactTitle
            DecalVendor(11) = Nz(RST!contactEmail, "-")
            DecalVendor(12) = Nz(RST!contactPrimaryPhone, "-")
        End If
    End If
Cleanup:
    RST.close
    Set RST = Nothing
End Sub

Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
'Purpose:   Return True if the file exists, even if it is hidden.
'Arguments: strFile: File name to look for. Current directory searched if no path included.
'           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
'Note:      Does not look inside subdirectories for the file.
'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory)    'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Function FolderExists(strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0 Then
        If Right(varIn, 1) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function



Public Function FormatExcel(ByVal TableNamex As String, sFile As String, ByVal sTabName As String, Optional ByVal sClose As String, _
                            Optional ByVal xScript As String, Optional ByVal sUnitCount As Integer, Optional WorkSheetTitleX As String, Optional ByVal LogoTrue As Boolean)
'_________________________
'XLS = Col letter,Format Option
' first set = number of columns
'Format Options TypeCommand
'Types:     T = Text | B = Bold | $ = Currency
'Commands:  C = Center | L=Lexf | R=Right | E = Sum col | X= Total Col Count |

    DoCmd.SetWarnings False
    Dim xlApp  As Object
    Dim xlSheet As Object
    Dim i      As Integer
    Set xlApp = CreateObject("Excel.Application")
    If sFile Like "*.xlsx" Then sFile = sFile Else sFile = sFile & ".xlsx"
    Set xlApp = CreateObject("Excel.Application")
    'Workbooks(sFile).close
    If FileExists(sFile) Then
        ' xlApp.Workbooks.close (sFile)
        On Error GoTo Err_ModifyExportedExcelFileFormats
        Kill sFile
    End If
    DoCmd.TransferSpreadsheet acExport, , TableNamex, sFile

    DoCmd.SetWarnings True
    On Error GoTo Err_ModifyExportedExcelFileFormats

    Application.SetOption "Show Status Bar", True
    vStatusBar = SysCmd(acSysCmdSetStatus, "Formatting export file... please wait.")
    Set xlSheet = xlApp.Workbooks.Open(sFile).Sheets(1)
    With xlApp
        .Application.Visible = True
        .Application.Sheets(TableNamex).Select
        .Application.ActiveSheet.Name = sTabName
        .Application.Cells.Select
        .Application.Selection.Font.Name = "Calibri"
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
        .Application.ActiveWorkbook.Save


        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit

        'Make Header:
        .Application.rows("1:1").Select
        .Application.Selection.Interior.ColorIndex = 25
        .Application.Selection.Font.Color = vbWhite


        If WorkSheetTitleX <> "" Then
            'Insert Client Name
            .Application.Range("B1").Select
            .Application.Selection.EntireRow.Insert
            .Application.Selection.EntireRow.Insert
            .Application.rows("1:1").Select
            .Application.Selection.Interior.ColorIndex = xlNone
            .Application.Range("B1:L1").Select
            .Application.Selection.Merge
            .Application.Selection.VerticalAlignment = xlCenter
            .Application.Selection.Font.Bold = True
            .Application.Selection.Font.Size = 20
            .Application.Selection.RowHeight = 20
            .Application.Range("B1") = WorkSheetTitleX
            .Application.Selection.rows.AutoFit
        End If
        If sUnitCount < 1 Then sUnitCount = .Application.Range("A" & .rows.Count).End(xlUp).Row - 3
        If LogoTrue = True Then
            'Insert Logo
            Set PICx = .Application.ActiveSheet.Shapes.AddPicture("C:\FAAS\Images\logo2.jpg", False, True, 3, 3, 100, 100)
            With PICx
                .ScaleHeight 0.5, True
                .ScaleWidth 0.5, True
            End With

        End If
        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit
        .Application.Range("A2").Select
        Dim ColXi As Integer

        If xScript & "test" <> "test" Then
            Dim TimeX, LoopX As Integer
            Dim Colz, InsZ, xi As String
            For i = 1 To Len(xScript)
                xi = Mid(xScript, i, 1)
                If xi = "|" Then TimeX = TimeX + 1
            Next i

            For LoopX = 1 To TimeX - 1
                Colz = Split(xScript, "|")(LoopX)
                InsZ = Split(Colz, ",")(1)
                Colz = Split(Colz, ",")(0)
                .Application.Range(Colz & "4" & ":" & Colz & (sUnitCount + 4)).Select
                If InsZ Like "*Sort*" And InsZ Like "*Desc*" Then .Application.Range(Colz & "3" & ":" & Colz & (sUnitCount + 3)).Sort Key1:=Range(Colz & "3"), Order1:=xlDescending, Header:=xlYes
                If InsZ Like "*Sort*" And InsZ Like "*Asc*" Then .Application.Range(Colz & "3" & ":" & Colz & (sUnitCount + 3)).Sort Key1:=Range(Colz & "3"), Order1:=xlAscending, Header:=xlYes
                If InsZ Like "$*" Then .Application.Selection.Style = "Currency"
                If InsZ Like "%*" Then .Application.Selection.Style = "Percent"
                If InsZ Like "S*" Then .Application.Selection.NumberFormat = "#,##"
                If InsZ Like "P1*" Then .Application.Selection.NumberFormat = "0.0%"
                If InsZ Like "P2*" Then .Application.Selection.NumberFormat = "0.00%"
                If InsZ Like "P3*" Then .Application.Selection.NumberFormat = "0.000%"
                If InsZ Like "*L" Then .Application.Selection.HorizontalAlignment = xlLeft
                If InsZ Like "*R" Then .Application.Selection.HorizontalAlignment = xlRight
                If InsZ Like "*C" Then .Application.Selection.HorizontalAlignment = xlCenter
                If InsZ Like "B*" Then .Application.Selection.Font.Bold = True
                If InsZ = "HH" Then .Application.Selection.EntireColumn.Hidden = True
                If InsZ Like "*X" Then .Application.Range(Colz & "2").formula = "=SUBTOTAL(103," & Colz & "4:" & Colz & (sUnitCount + 3) & ") & "" Items"""
                If InsZ Like "*E" Then
                    .Application.Range(Colz & "2").formula = "=Text(SUBTOTAL(109," & Colz & "4" & ":" & Colz & (sUnitCount + 3) & ")""#,#"")"
                    If InsZ Like "$*" Then .Application.Range(Colz & "2").Style = "Currency"
                    .Application.Range(Colz & "2").Interior.ColorIndex = 24
                    .Application.Range(Colz & "2:" & Colz & ((sUnitCount + 3))).Select
                    .Application.Selection.Columns.AutoFit
                    .Application.Range(Colz & "4:" & Colz & ((sUnitCount + 3))).Select
                    .Application.Selection.Interior.ColorIndex = 24
                            ' ?? Add border around A4
                    With .Range(Colz & "2").Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Range(Colz & "2").Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Range(Colz & "2").Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Range(Colz & "2").Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End If
                If .Application.Range(Colz & "2").Text Like "*-*" Then
                    .Application.ActiveCell.Offset(0, -1).Select
                    If IsEmpty(.Application.ActiveCell) Then
                        If .Application.Range(Colz & "3").Text Like "*Refrig*" Then
                            .Application.Columns((ColText2Num(Colz) - 2)).EntireColumn.Hidden = True
                        End If
                        .Application.Columns((ColText2Num(Colz) - 1)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 0)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 1)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 2)).EntireColumn.Hidden = True

                    End If
                End If
                If InsZ = "HY" Then
                    Dim RowX As Integer
                    Dim HyText, DisplayText, HyAddr As String
                    For RowX = 3 To (sUnitCount + 3)
                        HyText = .Application.Range(Colz & RowX).Value
                        If HyText Like "*Http*" Then
                            If HyText Like "*#*" Then
                                DisplayText = Split(HyText, "#")(0)
                                HyAddr = Split(HyText, "#")(1)
                                HyAddr = Split(HyAddr, "#")(0)
                            Else
                                HyAddr = HyText
                                DisplayText = "Web Link"
                            End If
                            .Application.Range(Colz & RowX).Select
                            .Application.Selection.Hyperlinks.Add _
                                    Anchor:=.Application.Range(Colz & RowX), _
                                    Address:=HyAddr, _
                                    ScreenTip:="Click to Open Hyperlink", _
                                    TextToDisplay:=DisplayText
                        End If
                    Next RowX
                    .Application.Range(Colz & 3).Select
                    .Application.Selection.Columns.AutoFit
                End If
                If InsZ = "DD" Then
                    .Application.rows("3:3").Select
                    .Application.Selection.RowHeight = 27
                    .Application.Selection.VerticalAlignment = xlTop
                    .Application.Selection.HorizontalAlignment = xlCenter
                    .Application.Selection.WrapText = True
                End If
                If InsZ = "D3" Then
                    .Application.rows("3:3").Select
                    .Application.Selection.RowHeight = 40
                    .Application.Selection.VerticalAlignment = xlTop
                    .Application.Selection.HorizontalAlignment = xlCenter
                    .Application.Selection.WrapText = True
                End If
                If InsZ Like "W*" Then .Application.Selection.ColumnWidth = Split(InsZ, "W")(1)
                If InsZ Like "DisplayAll" Then
                    .Application.Columns().Select
                    .Application.Selection.Columns.AutoFit
                End If
MoveOn:
            Next LoopX
        End If
        .Application.Range("B4").Select
CloseMe:
        .Application.ActiveWorkbook.Save
        If sClose = "Close" Then
            .Application.ActiveWorkbook.close
            .Quit
        End If
    End With
Cleanup:

    Set xlApp = Nothing
    Set xlSheet = Nothing

    vStatusBar = SysCmd(acSysCmdClearStatus)

Exit_ModifyExportedExcelFileFormats:
    Exit Function

Err_ModifyExportedExcelFileFormats:
    vStatusBar = SysCmd(acSysCmdClearStatus)
    If Err.Number = 70 Then
        MsgBox Err.Description & ": The Excel File (" & sFile & ") is already open. Please Close it or Save it as Different Filename", , "FAAS System"
    Else
 '       MsgBox Err.Number & " - " & Err.Description, , "FAAS System"
    End If
    Resume Exit_ModifyExportedExcelFileFormats

End Function

Public Function ColText2Num(ByVal CollLetters As String)
    CollLetters = UCase(CollLetters)
    If Len(CollLetters) = 1 Then ColText2Num = Asc(Mid$(CollLetters, 1, 1)) - 64
    If Len(CollLetters) = 2 Then ColText2Num = 26 + (Asc(Mid$(CollLetters, 2, 1)) - 64)
    If Len(CollLetters) = 2 And Left(CollLetters, 1) = "B" Then ColText2Num = 52 + (Asc(Mid$(CollLetters, 2, 1)) - 64)
    If Len(CollLetters) = 2 And Left(CollLetters, 1) = "C" Then ColText2Num = 78 + (Asc(Mid$(CollLetters, 2, 1)) - 64)
    If Len(CollLetters) = 3 Then ColText2Num = 52 + (Asc(Mid$(CollLetters, 2, 1)) - 64)

End Function
Public Function ExcelCustodian1(ByVal TableNamex As String, sFile As String, ByVal sTabName As String, Optional ByVal sClose As String, _
                                Optional ByVal xScript As String, Optional ByVal sUnitCount As Integer, Optional WorkSheetTitleX As String, Optional ByVal LogoTrue As Boolean, Optional HDRAligment As String)
'_________________________
'XLS = Col letter,Format Option
' first set = number of columns
'Format Options TypeCommand
'Types:     T = Text | B = Bold | $ = Currency
'Commands:  C = Center | L=Left | R=Right | E = Sum col | X= Total Count col |

    DoCmd.SetWarnings False
    Dim xlApp As Object
    Dim xlSheet As Object
    Dim i As Integer
    Set xlApp = CreateObject("Excel.Application")
    If sFile Like "*.xlsx" Then sFile = sFile Else sFile = sFile & ".xlsx"
    Set xlApp = CreateObject("Excel.Application")
    'Workbooks(sFile).close
    If FileExists(sFile) Then
        ' xlApp.Workbooks.close (sFile)
        On Error GoTo Err_ModifyExportedExcelFileFormats
        Kill sFile
    End If
    DoCmd.TransferSpreadsheet acExport, , TableNamex, sFile

    DoCmd.SetWarnings True
    On Error GoTo Err_ModifyExportedExcelFileFormats

    Application.SetOption "Show Status Bar", True
    vStatusBar = SysCmd(acSysCmdSetStatus, "Formatting export file... please wait.")
    Set xlSheet = xlApp.Workbooks.Open(sFile).Sheets(1)
    With xlApp
        .Application.Visible = True
        .Application.Sheets(TableNamex).Select
        .Application.ActiveSheet.Name = sTabName
        .Application.Cells.Select
        .Application.Selection.Font.Name = "Calibri"
        .Application.Selection.Font.Size = 10
        If Left(HDRAligment, 1) = "C" Then .Application.Selection.VerticalAlignment = xlCenter
        If Right(HDRAligment, 1) = "C" Then .Application.Selection.HorizontalAlignment = xlCenter
        If Not IsNull(HDRAligment) Then .Application.Selection.WrapText = True
        .Application.rows("1:1").Select
        If Left(HDRAligment, 1) = "T" Then .Application.Selection.VerticalAlignment = xlTop
        If Left(HDRAligment, 1) = "B" Then .Application.Selection.VerticalAlignment = xlBottom
        If Left(HDRAligment, 1) = "C" Then .Application.Selection.VerticalAlignment = xlCenter
        If Right(HDRAligment, 1) = "C" Then .Application.Selection.HorizontalAlignment = xlCenter
        If Right(HDRAligment, 1) = "L" Then .Application.Selection.HorizontalAlignment = xlLeft
        If Right(HDRAligment, 1) = "R" Then .Application.Selection.HorizontalAlignment = xlRight

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
        .Application.ActiveWorkbook.Save


        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit

        'Make Header:
        .Application.rows("1:1").Select

        If WorkSheetTitleX <> "" Then
            'Insert Title / Name
            .Application.Range("B1").Select
            .Application.Selection.EntireRow.Insert
            .Application.Selection.EntireRow.Insert
            .Application.rows("1:1").Select
            .Application.Selection.Interior.ColorIndex = xlNone
            .Application.Range("B1:L1").Select
            .Application.Selection.Merge
            .Application.Selection.VerticalAlignment = xlCenter
            .Application.Selection.Font.Bold = True
            .Application.Selection.Font.Size = 20
            .Application.Selection.RowHeight = 20
            .Application.Range("B1") = WorkSheetTitleX
            .Application.Selection.rows.AutoFit
        End If
        If LogoTrue = True Then
            'Insert Logo
            Set PICx = .Application.ActiveSheet.Shapes.AddPicture("C:\FAAS\Images\logo2.jpg", False, True, 3, 3, 100, 100)
            With PICx
                .ScaleHeight 0.5, True
                .ScaleWidth 0.5, True
            End With

        End If
        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit
        .Application.Range("A2").Select
        Dim ColXi As Integer

        If xScript & "test" <> "test" Then
            Dim TimeX, LoopX, Rowi As Integer
            Dim Colz, InsZ, RefColz, xi As String
            For i = 1 To Len(xScript)
                xi = Mid(xScript, i, 1)
                If xi = "|" Then TimeX = TimeX + 1
            Next i

            For LoopX = 1 To TimeX - 1
                Colz = Split(xScript, "|")(LoopX)
                InsZ = Split(Colz, ",")(1)
                Colz = Split(Colz, ",")(0)
                .Application.Range(Colz & "4" & ":" & Colz & (sUnitCount + 4)).Select
                If InsZ Like "*Sort*" And InsZ Like "*Desc*" Then .Application.Range(Colz & "3" & ":" & Colz & (sUnitCount + 3)).Sort Key1:=Range(Colz & "3"), Order1:=xlDescending, Header:=xlYes
                If InsZ Like "*Sort*" And InsZ Like "*Asc*" Then .Application.Range(Colz & "3" & ":" & Colz & (sUnitCount + 3)).Sort Key1:=Range(Colz & "3"), Order1:=xlAscending, Header:=xlYes
                If InsZ Like "$*" Then .Application.Selection.Style = "Currency"
                If InsZ Like "%*" Then .Application.Selection.Style = "Percent"
                If InsZ Like "S*" Then .Application.Selection.NumberFormat = "#,##"
                If InsZ Like "P1*" Then .Application.Selection.NumberFormat = "0.0%"
                If InsZ Like "P2*" Then .Application.Selection.NumberFormat = "0.00%"
                If InsZ Like "P3*" Then .Application.Selection.NumberFormat = "0.000%"
                If InsZ Like "*L" Then .Application.Selection.HorizontalAlignment = xlLeft
                If InsZ Like "*R" Then .Application.Selection.HorizontalAlignment = xlRight
                If InsZ Like "*C" Then .Application.Selection.HorizontalAlignment = xlCenter
                If InsZ Like "B*" Then .Application.Selection.Font.Bold = True
                If InsZ = "HH" Then .Application.Selection.EntireColumn.Hidden = True
                If InsZ Like "*X" Then .Application.Range(Colz & "2").formula = "=TEXT(SUBTOTAL(103," & Colz & "4:" & Colz & (sUnitCount + 3) & ")""#,#"") & "" Items """
                If InsZ Like "*E" Then
                    .Application.Range(Colz & "2").formula = "=Text(SUBTOTAL(109," & Colz & "4" & ":" & Colz & (sUnitCount + 3) & ")""#,#"""
                    If InsZ Like "$*" Then .Application.Range(Colz & "2").Style = "Currency"
                    .Application.Range(Colz & "2").Interior.ColorIndex = 24
                    .Application.Range(Colz & "2:" & Colz & ((sUnitCount + 3))).Select
                    .Application.Selection.Columns.AutoFit
                    .Application.Range(Colz & "4:" & Colz & ((sUnitCount + 3))).Select
                    .Application.Selection.Interior.ColorIndex = 24
                End If
                If .Application.Range(Colz & "2").Value = 0 Then
                    .Application.ActiveCell.Offset(0, -1).Select
                    If IsEmpty(.Application.ActiveCell) Then
                        .Application.Columns((ColText2Num(Colz) - 2)).Select
                        .Application.Columns((ColText2Num(Colz) - 1)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 0)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 1)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 2)).EntireColumn.Hidden = True
                    End If
                End If
            Next LoopX
        End If
        .Application.Range("A3").Select
        .Application.Selection.RowHeight = 40

        .Application.Range("B2").Select
        .Application.Selection.RowHeight = 25

        .Application.Range("C2").Select

        .Application.Range("D2").Select
        .Application.Selection.Font.Color = vbBlue
        .Application.Selection.Font.Bold = True
        .Application.Range("D2") = "FILE TYPE" & vbNewLine & "(Adress)"

        .Application.Range("E2").Select
        .Application.Range("E3") = "Date Equipment" & vbNewLine & "File Sent" & vbNewLine & " To Custodian"

        .Application.Range("F2").Select
        .Application.Range("F3") = "Equipment File" & vbNewLine & "Identifier 1" & vbNewLine & "(Title/Document#)"

        .Application.Range("G2").Select
        .Application.Range("G3") = "Equipment File" & vbNewLine & "Identifier 2" & vbNewLine & "(State of Title)"

        .Application.Range("H2").Select
        .Application.Selection.Font.Color = vbBlue
        .Application.Selection.Font.Bold = True
        .Application.Range("H2") = "OBLIGOR NAME" & vbNewLine & "(Borrower Name)"
        .Application.Range("H3") = "Equipment File" & vbNewLine & "Identifier 3" & vbNewLine & "(VIN)"

        .Application.Range("J2").Select
        .Application.Selection.Font.Color = vbBlue
        .Application.Selection.Font.Bold = True
        .Application.Range("J2") = "SUB-TRUST IDENTIFIER" & vbNewLine & "(Alt ID)"

        .Application.Range("K2").Select
        .Application.Selection.Font.Color = vbBlue
        .Application.Selection.Font.Bold = True
        .Application.Range("K2") = "LENDER IDENTIFIER" & vbNewLine & "(Trust #)"

        .Application.Range("N2").Select
        .Application.Selection.Font.Color = vbBlue
        .Application.Selection.Font.Bold = True
        .Application.Range("N2") = "ACCOUNT NO." & vbNewLine & "(COLL_ID)"

        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit
        .Application.Range("L2").Select
        .Application.Selection.ColumnWidth = 16
        .Application.Range("M2").Select
        .Application.Selection.ColumnWidth = 16
        .Application.Range("P2").Select
        .Application.Selection.EntireColumn.Delete
        .Application.Range("O2").Select
        .Application.Selection.EntireColumn.Delete
        .Application.Range("A2").Select
SetMargins:

  '   Dim LR As Long
   '  LR = Range("A" & Rows.Count).End(xlUp).Row
    .Application.ActiveSheet.PageSetup.PaperSize = xlPaperLegal
    .Application.ActiveSheet.PageSetup.PrintArea = "A1:N" & sUnitCount + 4
    .Application.ActiveSheet.PageSetup.LeftMargin = .Application.InchesToPoints(0.1)
    .Application.ActiveSheet.PageSetup.RightMargin = .Application.InchesToPoints(0.1)
    .Application.ActiveSheet.PageSetup.TopMargin = .Application.InchesToPoints(0.3)
    .Application.ActiveSheet.PageSetup.BottomMargin = .Application.InchesToPoints(0.3)
    .Application.ActiveSheet.PageSetup.Orientation = xlLandscape
    .Application.ActiveSheet.PageSetup.Zoom = False
    .Application.ActiveSheet.PageSetup.FitToPagesWide = 1
    .Application.ActiveSheet.PageSetup.FitToPagesTall = 10
CloseMe:
        .Application.ActiveWorkbook.Save
        If sClose = "Close" Then
            .Application.ActiveWorkbook.close
            .Quit
        End If
    End With
Cleanup:

    Set xlApp = Nothing
    Set xlSheet = Nothing

    vStatusBar = SysCmd(acSysCmdClearStatus)

Exit_ModifyExportedExcelFileFormats:
    Exit Function

Err_ModifyExportedExcelFileFormats:
    vStatusBar = SysCmd(acSysCmdClearStatus)
    If Err.Number = 70 Then
        MsgBox Err.Description & ": The Excel File (" & sFile & ") is already open. Please Close it or Save it as Different Filename", , "FAAS System"
    Else
        MsgBox Err.Number & " - " & Err.Description, , "FAAS System"
    End If
    Resume Exit_ModifyExportedExcelFileFormats

End Function
Public Function ExcelCustodian2(ByVal TableNamex As String, sFile As String, ByVal sTabName As String, Optional ByVal sClose As String, _
                                Optional ByVal xScript As String, Optional ByVal sUnitCount As Integer, Optional WorkSheetTitleX As String, Optional ByVal LogoTrue As Boolean, Optional HDRAligment As String)

    DoCmd.SetWarnings False
    Dim xlApp As Object
    Dim xlSheet As Object
    Dim i As Integer
    Set xlApp = CreateObject("Excel.Application")
    If sFile Like "*.xlsx" Then sFile = sFile Else sFile = sFile & ".xlsx"
    Set xlApp = CreateObject("Excel.Application")
    'Workbooks(sFile).close
    If FileExists(sFile) Then
   '     xlApp.Workbooks.close (sFile)
        On Error GoTo Err_ModifyExportedExcelFileFormats
        Kill sFile
    End If
    DoCmd.TransferSpreadsheet acExport, , TableNamex, sFile

    DoCmd.SetWarnings True
  '  On Error GoTo Err_ModifyExportedExcelFileFormats

    Application.SetOption "Show Status Bar", True
    vStatusBar = SysCmd(acSysCmdSetStatus, "Formatting export file... please wait.")
    Set xlSheet = xlApp.Workbooks.Open(sFile).Sheets(1)
    With xlApp
        .Application.Visible = True
        .Application.Sheets(TableNamex).Select
        .Application.ActiveSheet.Name = sTabName
        .Application.Cells.Select
        '  .Application.Selection.ClearFormats
        .Application.Selection.Font.Name = "Calibri"
        .Application.Selection.Font.Size = 10
        If Left(HDRAligment, 1) = "C" Then .Application.Selection.VerticalAlignment = xlCenter
        If Right(HDRAligment, 1) = "C" Then .Application.Selection.HorizontalAlignment = xlCenter
        If Not IsNull(HDRAligment) Then .Application.Selection.WrapText = True
        .Application.rows("1:1").Select
        If Left(HDRAligment, 1) = "T" Then .Application.Selection.VerticalAlignment = xlTop
        If Left(HDRAligment, 1) = "B" Then .Application.Selection.VerticalAlignment = xlBottom
        If Left(HDRAligment, 1) = "C" Then .Application.Selection.VerticalAlignment = xlCenter
        If Right(HDRAligment, 1) = "C" Then .Application.Selection.HorizontalAlignment = xlCenter
        If Right(HDRAligment, 1) = "L" Then .Application.Selection.HorizontalAlignment = xlLeft
        If Right(HDRAligment, 1) = "R" Then .Application.Selection.HorizontalAlignment = xlRight

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
        .Application.ActiveWorkbook.Save

        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit

        'Make Header:
        .Application.rows("1:1").Select
        If WorkSheetTitleX <> "" Then
            'Insert Title / Name
            .Application.Range("B1").Select
            .Application.Selection.EntireRow.Insert
            .Application.Selection.EntireRow.Insert
            .Application.rows("1:1").Select
            .Application.Selection.Interior.ColorIndex = xlNone
            .Application.Range("B1:E1").Select
            .Application.Selection.Merge
            .Application.Selection.VerticalAlignment = xlCenter
            .Application.Selection.Font.Bold = True
            .Application.Selection.Font.Size = 20
            .Application.Selection.RowHeight = 20
            .Application.Range("B1") = WorkSheetTitleX
            .Application.Selection.rows.AutoFit
        End If
        If LogoTrue = True Then
            'Insert Logo
            Set PICx = .Application.ActiveSheet.Shapes.AddPicture("C:\FAAS\Images\logo2.jpg", False, True, 3, 3, 100, 100)
            With PICx
                .ScaleHeight 0.5, True
                .ScaleWidth 0.5, True
            End With

        End If
        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit
        .Application.Range("A2").Select
        Dim ColXi As Integer

        If xScript & "test" <> "test" Then
            Dim TimeX, LoopX As Integer
            Dim Colz, InsZ, xi As String
            For i = 1 To Len(xScript)
                xi = Mid(xScript, i, 1)
                If xi = "|" Then TimeX = TimeX + 1
            Next i

            For LoopX = 1 To TimeX - 1
                Colz = Split(xScript, "|")(LoopX)
                InsZ = Split(Colz, ",")(1)
                Colz = Split(Colz, ",")(0)
                .Application.Range(Colz & "4" & ":" & Colz & (sUnitCount + 4)).Select

                If InsZ Like "*Sort*" And InsZ Like "*Desc*" Then .Application.Range(Colz & "3" & ":" & Colz & (sUnitCount + 3)).Sort Key1:=Range(Colz & "3"), Order1:=xlDescending, Header:=xlYes
                If InsZ Like "*Sort*" And InsZ Like "*Asc*" Then .Application.Range(Colz & "3" & ":" & Colz & (sUnitCount + 3)).Sort Key1:=Range(Colz & "3"), Order1:=xlAscending, Header:=xlYes
                If InsZ Like "$*" Then .Application.Selection.Style = "Currency"
                If InsZ Like "%*" Then .Application.Selection.Style = "Percent"
                If InsZ Like "S*" Then .Application.Selection.NumberFormat = "#,##"
                If InsZ Like "P1*" Then .Application.Selection.NumberFormat = "0.0%"
                If InsZ Like "P2*" Then .Application.Selection.NumberFormat = "0.00%"
                If InsZ Like "P3*" Then .Application.Selection.NumberFormat = "0.000%"
                If InsZ Like "*L" Then .Application.Selection.HorizontalAlignment = xlLeft
                If InsZ Like "*R" Then .Application.Selection.HorizontalAlignment = xlRight
                If InsZ Like "*C" Then .Application.Selection.HorizontalAlignment = xlCenter
                If InsZ Like "B*" Then .Application.Selection.Font.Bold = True
                If InsZ = "HH" Then .Application.Selection.EntireColumn.Hidden = True
                If InsZ Like "*X" Then .Application.Range(Colz & "2").formula = "=SUBTOTAL(103," & Colz & "4:" & Colz & (sUnitCount + 3) & ") & "" Items """
                If InsZ Like "*E" Then
                    .Application.Range(Colz & "2").formula = "=SUBTOTAL(109," & Colz & "4" & ":" & Colz & (sUnitCount + 3) & ")"
                    If InsZ Like "$*" Then .Application.Range(Colz & "2").Style = "Currency"
                    .Application.Range(Colz & "2").Interior.ColorIndex = 24
                    .Application.Range(Colz & "2:" & Colz & ((sUnitCount + 3))).Select
                    .Application.Selection.Columns.AutoFit
                    .Application.Range(Colz & "4:" & Colz & ((sUnitCount + 3))).Select
                    .Application.Selection.Interior.ColorIndex = 24
                End If
                If .Application.Range(Colz & "2").Value = 0 Then
                    .Application.ActiveCell.Offset(0, -1).Select
                    If IsEmpty(.Application.ActiveCell) Then
                        .Application.Columns((ColText2Num(Colz) - 2)).Select
                        .Application.Columns((ColText2Num(Colz) - 1)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 0)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 1)).EntireColumn.Hidden = True
                        .Application.Columns((ColText2Num(Colz) + 2)).EntireColumn.Hidden = True
                    End If
                End If
            Next LoopX
        End If
        .Application.Range("A3:E3").Select
        .Application.Selection.Font.Color = vbBlue
        .Application.Selection.Font.Bold = True
        .Application.Selection.RowHeight = 44

        .Application.Range("A3").Select
        .Application.Range("A3") = "ACCOUNT NO." & vbNewLine & "(COLL_ID)"

        .Application.Range("B3").Select
        .Application.Range("B3") = "OBLIGOR NAME (Borrower Name)"

        .Application.Range("C3").Select
        .Application.Range("C3") = "FILE TYPE " & vbNewLine & "(Address)"

        .Application.Range("D3").Select
        .Application.Range("D3") = "SUB-TRUST IDENTIFIER (Alt ID)"

        .Application.Range("E3").Select
        .Application.Range("E3") = "LENDER IDENTIFIER" & vbNewLine & "(Trust #)"

        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit
        .Application.Range("A2").Select

SetMargins:

  '   Dim LR As Long
   '  LR = Range("A" & Rows.Count).End(xlUp).Row
    .Application.ActiveSheet.PageSetup.PrintArea = "A1:E" & sUnitCount + 4
    .Application.ActiveSheet.PageSetup.LeftMargin = .Application.InchesToPoints(0.1)
    .Application.ActiveSheet.PageSetup.RightMargin = .Application.InchesToPoints(0.1)
    .Application.ActiveSheet.PageSetup.TopMargin = .Application.InchesToPoints(0.3)
    .Application.ActiveSheet.PageSetup.BottomMargin = .Application.InchesToPoints(0.3)
    .Application.ActiveSheet.PageSetup.Orientation = xlLandscape
    .Application.ActiveSheet.PageSetup.Zoom = False
    .Application.ActiveSheet.PageSetup.FitToPagesWide = 1
    .Application.ActiveSheet.PageSetup.FitToPagesTall = 10


CloseMe:
        .Application.ActiveWorkbook.Save
        If sClose = "Close" Then
            .Application.ActiveWorkbook.close
            .Quit
        End If
    End With
Cleanup:

    Set xlApp = Nothing
    Set xlSheet = Nothing

    vStatusBar = SysCmd(acSysCmdClearStatus)

Exit_ModifyExportedExcelFileFormats:
    Exit Function

Err_ModifyExportedExcelFileFormats:
    vStatusBar = SysCmd(acSysCmdClearStatus)
    If Err.Number = 70 Then
        MsgBox Err.Description & ": The Excel File (" & sFile & ") is already open. Please Close it or Save it as Different Filename", , "FAAS System"
    Else
        MsgBox Err.Number & " - " & Err.Description, , "FAAS System"
    End If
    Resume Exit_ModifyExportedExcelFileFormats

End Function
Public Function SendFormData2Excel(frm As Form, ByVal sTabName As String, Optional ByVal WorkSheetTitleX As String, Optional ByVal LogoTrue As Boolean, Optional ByVal xScript As String, Optional ByVal sUnitCount As Integer)
    ' frm is the name of the form you want to send to Excel
    ' strSheetName is the name of the sheet you want to name it to

    'Dim Rst As DAO.Recordset
    Dim xlApp  As Object
    Dim xlWBk  As Object
    Dim xlSheet As Object
    Dim intCount As Integer
    Const xlCenter As Long = -4108
    Const xlBottom As Long = -4107

    On Error GoTo Err_Handler

    Set RST = frm.RecordsetClone

    Set xlApp = CreateObject("Excel.Application")
    Set xlWBk = xlApp.Workbooks.Add
    xlApp.Visible = True

    Set xlSheet = xlWBk.Worksheets("Sheet1")

    If Len(WorkSheetTitleX) > 0 Then
        xlSheet.Name = Left(sTabName, 34)
    End If
    xlSheet.Range("A1").Select
    Do Until intCount = RST.Fields.Count
        xlApp.ActiveCell = RST.Fields(intCount).Name
        xlApp.ActiveCell.Offset(0, 1).Select
        intCount = intCount + 1
    Loop
    RST.MoveFirst
    xlSheet.Range("A2").CopyFromRecordset RST
    xlSheet.Range("1:1").Select

    RST.close
    Set RST = Nothing
    With xlApp
        .Application.Sheets(xlSheet.Name).Select
        .Application.ActiveSheet.Name = sTabName
        .Application.Cells.Select
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

        .Application.Columns().Select
        .Application.Selection.Columns.AutoFit

        'Make Header:
        .Application.rows("1:1").Select
        .Application.Selection.Interior.ColorIndex = 25
        .Application.Selection.Font.Color = vbWhite


        If WorkSheetTitleX <> "" Then
            'Insert Client Name
            .Application.Range("B1").Select
            .Application.Selection.EntireRow.Insert
            .Application.Selection.EntireRow.Insert
            .Application.rows("1:1").Select
            .Application.Selection.Interior.ColorIndex = xlNone
            .Application.Range("B1:L1").Select
            .Application.Selection.Merge
            .Application.Selection.VerticalAlignment = xlCenter
            .Application.Selection.Font.Bold = True
            .Application.Selection.Font.Size = 20
            .Application.Selection.RowHeight = 22
            .Application.Range("B1") = WorkSheetTitleX
        End If
        If sUnitCount < 1 Then sUnitCount = .Application.Range("A" & .rows.Count).End(xlUp).Row - 3
        If LogoTrue = True Then
            'Insert Logo
            Set PICx = .Application.ActiveSheet.Shapes.AddPicture("C:\FAAS\Images\logo2.jpg", False, True, 3, 3, 100, 100)
            With PICx
                .ScaleHeight 0.5, True
                .ScaleWidth 0.5, True
            End With
        End If
        If xScript & "test" <> "test" Then
            Dim TimeX, LoopX As Integer
            Dim Colz, InsZ, xi As String
            For i = 1 To Len(xScript)
                xi = Mid(xScript, i, 1)
                If xi = "|" Then TimeX = TimeX + 1
            Next i
            'XLS = Col letter,Format Option
            ' first set = number of columns
            'Format Options TypeCommand
            'Types:     T = Text | B = Bold | $ = Currency
            'Commands:  C = Center | L=Lexf | R=Right | E = Sum col | X= Total Col Count |
            For LoopX = 1 To TimeX - 1
                Colz = Split(xScript, "|")(LoopX)
                InsZ = Split(Colz, ",")(1)
                Colz = Split(Colz, ",")(0)
                .Application.Range(Colz & "4" & ":" & Colz & (sUnitCount + 4)).Select

                If InsZ Like "*Sort*" And InsZ Like "*Desc*" Then .Application.Range(Colz & "3" & ":" & Colz & (sUnitCount + 3)).Sort Key1:=Range(Colz & "3"), Order1:=xlDescending, Header:=xlYes
                If InsZ Like "*Sort*" And InsZ Like "*Asc*" Then .Application.Range(Colz & "3" & ":" & Colz & (sUnitCount + 3)).Sort Key1:=Range(Colz & "3"), Order1:=xlAscending, Header:=xlYes
                If InsZ Like "$*" Then .Application.Selection.Style = "Currency"
                If InsZ Like "%*" Then .Application.Selection.Style = "Percent"
                If InsZ Like "S*" Then .Application.Selection.NumberFormat = "#,##0"
                If InsZ Like "P1*" Then .Application.Selection.NumberFormat = "0.0%"
                If InsZ Like "P2*" Then .Application.Selection.NumberFormat = "0.00%"
                If InsZ Like "P3*" Then .Application.Selection.NumberFormat = "0.000%"
                If InsZ Like "P4*" Then .Application.Selection.NumberFormat = "0.0000%"
                If InsZ Like "P5*" Then .Application.Selection.NumberFormat = "0.00000%"
                If InsZ Like "*L" Then .Application.Selection.HorizontalAlignment = xlLeft
                If InsZ Like "*R" Then .Application.Selection.HorizontalAlignment = xlRight
                If InsZ Like "*C" Then .Application.Selection.HorizontalAlignment = xlCenter
                If InsZ Like "B*" Then .Application.Selection.Font.Bold = True
                If InsZ = "HH" Then .Application.Selection.EntireColumn.Hidden = True
                If InsZ = "ZZ" Then .Application.Selection.EntireColumn.Delete
                If InsZ Like "*X" Then .Application.Range(Colz & "2").formula = "=NUMBERVALUE(SUBTOTAL(103," & Colz & "4:" & Colz & (sUnitCount + 3) & "),""#,#"") & "" Items """
                 If InsZ Like "A*" Then .Application.Range(Colz & "2").formula = "=TEXT(Round(SUBTOTAL(101," & Colz & "4" & ":" & Colz & (sUnitCount + 3) & "),0),""#,#"")"
                If InsZ Like "A*" Then .Application.Range(Colz & "2").HorizontalAlignment = xlRight
                If InsZ Like "*E" Then
                     If InsZ Like "$*" Then .Application.Range(Colz & "2").formula = "=NUMBERVALUE(SUBTOTAL(109," & Colz & "4" & ":" & Colz & (sUnitCount + 3) & "),""$#,##0.00"")" _
                                       Else .Application.Range(Colz & "2").formula = "=NUMBERVALUE(SUBTOTAL(109," & Colz & "4" & ":" & Colz & (sUnitCount + 3) & "),""#,#"")"

                    If InsZ Like "$*" Then .Application.Range(Colz & "2").Style = "Currency"
                    .Application.Range(Colz & "2").Interior.ColorIndex = 24
                    .Application.Range(Colz & "2:" & Colz & ((sUnitCount + 3))).Select
                    .Application.Selection.Columns.AutoFit
                    .Application.Range(Colz & "4:" & Colz & ((sUnitCount + 3))).Select

                    If .Application.Range(Colz & "2").Value = 0 Then
                        .Application.ActiveCell.Offset(0, -1).Select
                        .Application.Columns((ColText2Num(Colz) + 0)).EntireColumn.Hidden = True
                    End If

                End If
                If InsZ = "HY" Then
                    Dim RowX As Integer
                    Dim HyText, DisplayText, HyAddr As String
                    For RowX = 3 To (sUnitCount + 3)
                        HyText = .Application.Range(Colz & RowX).Value
                        If HyText Like "*Http*" Then
                            If HyText Like "*#*" Then
                                DisplayText = Split(HyText, "#")(0)
                                HyAddr = Split(HyText, "#")(1)
                                HyAddr = Split(HyAddr, "#")(0)
                            Else
                                HyAddr = HyText
                                DisplayText = "Web Link"
                            End If
                            .Application.Range(Colz & RowX).Select
                            .Application.Selection.Hyperlinks.Add _
                                Anchor:=.Application.Range(Colz & RowX), _
                                Address:=HyAddr, _
                                ScreenTip:="Click to Open Hyperlink", _
                                TextToDisplay:=DisplayText
                        End If
                    Next RowX
                    .Application.Range(Colz & 3).Select
                    .Application.Selection.Columns.AutoFit
                End If
                If InsZ = "DD" Then
                    .Application.rows("3:3").Select
                    .Application.Selection.RowHeight = 25
                    .Application.Selection.VerticalAlignment = xlTop
                    .Application.Selection.HorizontalAlignment = xlCenter
                    .Application.Selection.WrapText = True
                End If
                If InsZ Like "W*" Then .Application.Selection.ColumnWidth = Split(InsZ, "W")(1)
                If InsZ Like "DisplayAll" Then
                    .Application.Columns().Select
                    .Application.Selection.Columns.AutoFit
                End If
            Next LoopX
        End If
        .Application.rows("1:1").Select
    End With
    Exit Function
Err_Handler:
    DoCmd.SetWarnings True
    MsgBox Err.Description, vbExclamation, Err.Number
    Exit Function

End Function

Public Function SendReport2Excel(ByVal RSql As String, ByVal sTabName As String, Optional ByVal WorkSheetTitleX As String, Optional ByVal LogoTrue As Boolean, Optional ByVal xScript As String, Optional ByVal sUnitCount As Long, Optional ByVal ShowExcel As Boolean, Optional ByVal GroupReports As Integer, Optional ByVal LastReport As Boolean, Optional PrintOpts As String, Optional sFile As String)
    Dim rs As DAO.Recordset
    Dim xlApp As Object, xlWBk As Object, xlSheet As Object
    Dim SaveMe As Boolean
    Const xlCenter As Long = -4108
    Const xlBottom As Long = -4107
    SaveMe = (sFile <> "")
    If sFile = "" Then sFile = "C:\Faas\" & FaUserName & "\Temp\ExcelReport.xlsx"
    Call IfExcelOpenCloseIt(sFile)
    If IsOpen("MessagePopUp") = True Then MessageUser "Update", "Getting Ready", "Setting EXCEL Data..."
    Set rs = CurrentDb.OpenRecordset(RSql, dbOpenSnapshot)
    If rs.RecordCount < 1 Then Exit Function
    rs.MoveFirst
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = Not (IsNull(ShowExcel) Or ShowExcel = False)
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    If IsNull(GroupReports) Or GroupReports <= 1 Then
        If FileExists(sFile, 0) = True Then
            SetAttr sFile, vbNormal
            Kill sFile
        End If
        Set xlWBk = xlApp.Workbooks.Add
        Set xlSheet = xlWBk.Worksheets(1)
    Else
        Set xlWBk = xlApp.Workbooks.Open(sFile, True, False)
        xlWBk.Worksheets.Add After:=xlWBk.Worksheets(xlWBk.Worksheets.Count)
        Set xlSheet = xlWBk.ActiveSheet
    End If
    If Len(WorkSheetTitleX) > 0 Then xlSheet.Name = Left(sTabName, 34)
    Dim ColumnCount As Long
    ColumnCount = rs.Fields.Count
    Dim colIndex As Integer
    For colIndex = 0 To ColumnCount - 1
        xlSheet.Cells(1, colIndex + 1).Value = rs.Fields(colIndex).Name
    Next colIndex
    xlSheet.Range("A2").CopyFromRecordset rs
    rs.close
    Set rs = Nothing
    If IsOpen("MessagePopUp") = True Then MessageUser "Update", "Formatting Excel", "Setting Up the page"
    xlSheet.Name = sTabName
    xlSheet.Cells.Font.Name = "Calibri"
    xlSheet.Cells.Font.Size = 10
    xlSheet.rows(1).Font.Bold = True
    xlSheet.rows(1).Font.Size = 10
    xlSheet.Cells.RowHeight = 12.75
    xlSheet.Columns.AutoFit
    xlSheet.Range("A2").Select       ' Select the cell BELOW where you want the freeze line
    xlApp.ActiveWindow.FreezePanes = True
    xlSheet.Range("A1").AutoFilter
    xlSheet.Columns.AutoFit
    xlSheet.Range("A1:" & Num2ColumnLetter(ColumnCount) & "1").Interior.ColorIndex = 25
    xlSheet.Range("A1:" & Num2ColumnLetter(ColumnCount) & "1").Font.Color = vbWhite
    Dim titleInserted As Boolean
    titleInserted = (WorkSheetTitleX <> "")
    If titleInserted Then
        If IsOpen("MessagePopUp") = True Then MessageUser "Update", "Formatting Excel", "Setting Title"
        xlSheet.rows("1:1").Insert
        xlSheet.rows("1:1").Insert
        xlSheet.rows(1).Interior.ColorIndex = xlNone
        xlSheet.rows(1).VerticalAlignment = xlCenter
        xlSheet.rows(1).Font.Bold = True
        xlSheet.rows(1).Font.Size = 20
        xlSheet.rows(1).RowHeight = 20
        xlSheet.Range("B1").Value = WorkSheetTitleX
        xlSheet.Range("B1").ColumnWidth = 30
    End If
    If sUnitCount < 1 Then sUnitCount = xlSheet.Range("A" & xlSheet.rows.Count).End(xlUp).Row - IIf(titleInserted, 3, 1)
    If LogoTrue = True Then
        Dim PICx As Object
        If WorkSheetTitleX Like "*Factory*" Then
            Set PICx = xlSheet.Shapes.AddPicture("C:\FAAS\Images\FA_logo_RMKT.jpg", False, True, 1300, 0, 100, 100)
            With PICx
                .ScaleHeight 0.26, True
                .ScaleWidth 0.26, True
            End With
        Else
            Set PICx = xlSheet.Shapes.AddPicture("C:\FAAS\Images\logo2.jpg", False, True, 3, 3, 100, 100)
            With PICx
                .ScaleHeight 0.8, True
                .ScaleWidth 0.8, True
            End With
        End If
    End If
    If xScript & "test" <> "test" Then
        Dim TimeX As Integer, LoopX As Integer
        Dim Colz As String
        Dim InsZ As String
        Dim FormX As String, xi As String
        TimeX = 0
        For i = 1 To Len(xScript)
            xi = Mid(xScript, i, 1)
            If xi = "|" Then TimeX = TimeX + 1
        Next i
        If xScript Like "*FxHeader*" Then
            FormX = Split(xScript, "FxHeader")(1)
            FormX = Split(FormX, "|")(0)
        End If
        For LoopX = 1 To TimeX - 1
            Colz = Split(xScript, "|")(LoopX)
            If Split(xScript, "|")(LoopX) Like "*FxAll*" Then
                InsZ = "FxAll=" & Split(Colz, "FxAll=")(1)
            Else
                InsZ = Split(Colz, ",")(1)
            End If
            Colz = Split(Colz, ",")(0)
            If IsOpen("MessagePopUp") = True Then MessageUser "Update", "Tab: " & sTabName, "Formatting '" & Colz & "'-" & InsZ
            Dim dataStartRow As Long
            dataStartRow = IIf(titleInserted, 4, 2)
            Dim dataEndRow As Long
            dataEndRow = dataStartRow + sUnitCount - 1
            Dim sortRange As Object
            Set sortRange = xlSheet.Range(Colz & (dataStartRow - 1) & ":" & Colz & dataEndRow)
            If InsZ Like "*Sort*" And InsZ Like "*Desc*" Then sortRange.Sort Key1:=xlSheet.Range(Colz & (dataStartRow - 1)), Order1:=xlDescending, Header:=xlYes
            If InsZ Like "*Sort*" And InsZ Like "*Asc*" Then sortRange.Sort Key1:=xlSheet.Range(Colz & (dataStartRow - 1)), Order1:=xlAscending, Header:=xlYes
            Dim dataRange As Object
            Set dataRange = xlSheet.Range(Colz & dataStartRow & ":" & Colz & (dataEndRow + 1))
            If InsZ Like "$*" Then dataRange.Style = "Currency"
            If InsZ Like "$Int*" Then dataRange.NumberFormat = "$#,##0_);[Red]($#,##0)"
            If InsZ Like "S*" Then dataRange.NumberFormat = "#,##"
            If InsZ Like "SInt*" Then dataRange.NumberFormat = "#,##0_);[Red](#,##0)"
            If InsZ Like "MthYear*" Then dataRange.NumberFormat = "mmm-yyyy"
            If InsZ Like "%*" Then dataRange.Style = "Percent"
            If InsZ Like "P1*" Then dataRange.NumberFormat = "0.0%"
            If InsZ Like "P2*" Then dataRange.NumberFormat = "0.00%"
            If InsZ Like "P3*" Then dataRange.NumberFormat = "0.000%"
            If InsZ Like "P4*" Then dataRange.NumberFormat = "0.0000%"
            If InsZ Like "P5*" Then dataRange.NumberFormat = "0.00000%"
            If InsZ Like "P6*" Then dataRange.NumberFormat = "0.000000%"
            If InsZ Like "P7*" Then dataRange.NumberFormat = "0.0000000%"
            If InsZ Like "B*" Then dataRange.Font.Bold = True
            Dim totalRange As Object
            Set totalRange = xlSheet.Range(Colz & (dataStartRow - 2))
            If InsZ Like "A*" Then
                totalRange.formula = "=Text(Round(SUBTOTAL(101," & Colz & dataStartRow & ":" & Colz & dataEndRow & "),0),""#,#"")"
                totalRange.HorizontalAlignment = xlCenter
            End If
            If InsZ = "HH" Then xlSheet.Columns(Colz).Hidden = True
            If InsZ = "ZZ" Then xlSheet.Columns(Colz).Delete
            If Not InsZ Like "*cLabel*" And InsZ Like "*L" Then dataRange.HorizontalAlignment = xlLeft
            If Not InsZ Like "*cLabel*" And InsZ Like "*R" Then dataRange.HorizontalAlignment = xlRight
            If Not InsZ Like "*cLabel*" And InsZ Like "*C" Then dataRange.HorizontalAlignment = xlCenter
            If Not InsZ Like "*cLabel*" And InsZ Like "*X" Then totalRange.formula = "=Text(SUBTOTAL(103," & Colz & dataStartRow & ":" & Colz & dataEndRow & "),""#,#"") & "" Items """
            If Not InsZ Like "*:*" And Not InsZ Like "*cLabel*" And InsZ Like "*E" Then
                If InsZ Like "$*" Then
                    If InsZ Like "$Int*" Then
                        totalRange.formula = "=NUMBERVALUE(SUBTOTAL(109," & Colz & dataStartRow & ":" & Colz & dataEndRow & "),""$#,##0"")"
                        totalRange.NumberFormat = "$#,##0_);[Red]($#,##0)"
                    Else
                        totalRange.formula = "=NUMBERVALUE(SUBTOTAL(109," & Colz & dataStartRow & ":" & Colz & dataEndRow & "),""$#,##0.00"")"
                        totalRange.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
                    End If
                Else
                    totalRange.formula = "=NUMBERVALUE(SUBTOTAL(109," & Colz & dataStartRow & ":" & Colz & dataEndRow & "),""#,##0"")"
                    totalRange.NumberFormat = "#,##0_);[Red](#,##0)"
                End If
                totalRange.Interior.Color = RGB(204, 236, 255)
                xlSheet.Range(Colz & (dataStartRow - 2) & ":" & Colz & dataEndRow).Columns.AutoFit
                With totalRange.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With totalRange.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With totalRange.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With totalRange.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                If totalRange.Value = 0 Then xlSheet.Columns(ColText2Num(Colz)).Hidden = True
            End If
            If InsZ = "HY" Then
                Dim RowX As Integer
                Dim HyText As String, DisplayText As String, HyAddr As String
                For RowX = dataStartRow - 1 To dataEndRow
                    HyText = xlSheet.Range(Colz & RowX).Value
                    If HyText Like "*Http*" Then
                        If HyText Like "*#*" Then
                            DisplayText = Split(HyText, "#")(0)
                            HyAddr = Split(HyText, "#")(1)
                            HyAddr = Split(HyAddr, "#")(0)
                        Else
                            HyAddr = HyText
                            DisplayText = "Web Link"
                        End If
                        xlSheet.Hyperlinks.Add Anchor:=xlSheet.Range(Colz & RowX), Address:=HyAddr, ScreenTip:="Click to Open Hyperlink", TextToDisplay:=DisplayText
                    End If
                Next RowX
                xlSheet.Columns(Colz).AutoFit
            End If
            If InsZ = "Wrap" Then
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).WrapText = True
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).VerticalAlignment = xlTop
            End If
            If InsZ = "zTop" Then
                xlSheet.UsedRange.VerticalAlignment = xlTop
                xlSheet.rows.AutoFit
            End If
            If InsZ = "zCenter" Then
                xlSheet.UsedRange.VerticalAlignment = xlCenter
                xlSheet.rows.AutoFit
            End If
            If InsZ = "DD" Then
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).RowHeight = 25
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).VerticalAlignment = xlTop
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).HorizontalAlignment = xlCenter
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).WrapText = True
            End If
            If InsZ = "DD3" Then
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).RowHeight = 38
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).VerticalAlignment = xlTop
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).HorizontalAlignment = xlCenter
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).WrapText = True
            End If
            If InsZ = "DD4" Then
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).RowHeight = 52
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).VerticalAlignment = xlTop
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).HorizontalAlignment = xlCenter
                xlSheet.rows(dataStartRow - 1 & ":" & dataStartRow - 1).WrapText = True
            End If
            If InStr(InsZ, "mTable") > 0 Then
                Dim mStartCell As String, mStyle As String, mName As String
                Dim mRangeType As String, mTotalRow As Boolean
                Dim mLastRow As Long, mLastCol As Long
                Dim mRange As Object, mTable As Object
                mStartCell = "A3"
                mStyle = ""
                mName = ""
                mRangeType = "CurrentRegion"
                mTotalRow = False
                If InsZ Like "*Start=*" Then mStartCell = SplitString(SplitString(InsZ, "Start=", 1), ":", 0)
                If InsZ Like "*Use=*" Then mRangeType = LCase(SplitString(SplitString(InsZ, "Use=", 1), ":", 0))
                If InsZ Like "*Style=*" Then mStyle = "TableStyle" & SplitString(SplitString(InsZ, "Style=", 1), ":", 0)
                If InsZ Like "*Name=*" Then mName = SplitString(SplitString(InsZ, "Name=", 1), ":", 0)
                If InsZ Like "*TotalRow=*" Then mTotalRow = (LCase(SplitString(SplitString(InsZ, "TotalRow=", 1), ":", 0)) = "true")
                Select Case mRangeType
                    Case "currentregion"
                        Set mRange = xlSheet.Range(mStartCell).CurrentRegion
                    Case "usedrange"
                        Set mRange = xlSheet.UsedRange
                    Case "manual"
                        Dim mCol As String
                        mCol = Left(mStartCell, 1)
                        mLastRow = xlSheet.Cells(xlSheet.rows.Count, mCol).End(xlUp).Row
                        mLastCol = xlSheet.Cells(1, xlSheet.Columns.Count).End(xlToLeft).Column
                        Set mRange = xlSheet.Range(mStartCell, xlSheet.Cells(mLastRow, mLastCol))
                    Case Else
                        Set mRange = xlSheet.Range(mStartCell).CurrentRegion
                End Select
                Set mTable = xlSheet.ListObjects.Add(1, mRange, , 1)
                If InsZ Like "*ClearHeader*" Then
                    With mTable.HeaderRowRange.Interior
                        .Pattern = xlNone
                        .TintAndShade = 0
                        .ColorIndex = xlColorIndexNone
                    End With
                End If
                If mStyle <> "" Then mTable.TableStyle = mStyle
                If mName <> "" Then mTable.Name = mName
                mTable.showTotals = mTotalRow
            End If
            If InsZ = "LL1" Then
                With xlSheet.PageSetup
                    .CenterHorizontally = False
                    .CenterVertically = False
                    .Orientation = xlLandscape
                    .Draft = False
                    .PaperSize = xlPaperLetter
                    .FirstPageNumber = xlAutomatic
                    .Order = xlDownThenOver
                    .BlackAndWhite = False
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                End With
            End If
            If InsZ Like "LL*" Then
                xlSheet.Columns.AutoFit
                With xlSheet.PageSetup
                    .CenterHorizontally = False
                    .CenterVertically = False
                    .Orientation = xlLandscape
                    .Draft = False
                    If InsZ = "LL2" Then .PaperSize = xlPaperLetter
                    On Error Resume Next
                    If InsZ = "LLG" Then .PaperSize = xlPaperLegal
                    If InsZ = "LLG" And .PaperSize = 1 Then .PaperSize = xlPaperLetter
                    If InsZ = "LLT" Then .PaperSize = xlPaperTabloid
                    If InsZ = "LLT" And .PaperSize < 3 Then .PaperSize = xlPaperLegal
                    On Error GoTo 0
                    .FirstPageNumber = xlAutomatic
                    .Order = xlDownThenOver
                    .BlackAndWhite = False
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                End With
            End If
            Dim EndColz As String, RowZ As String, Textz As String, ColorZ As String
            If InsZ = "xlTop" Then xlSheet.UsedRange.VerticalAlignment = xlTop
            If InsZ = "xlCenter" Then xlSheet.UsedRange.VerticalAlignment = xlCenter
            If InsZ = "VAutoFit" Then xlSheet.rows.AutoFit
            If InsZ Like "*xlBorder*" Then
                EndColz = SplitString(InsZ, ":", 2)
                RowZ = SplitString(InsZ, ":", 1)
                If RowZ = 0 Then RowZ = sUnitCount + 4
                Dim borderRange As Object
                Set borderRange = xlSheet.Range(Colz & dataStartRow & ":" & EndColz & (dataStartRow + val(RowZ) - 1))
                borderRange.Borders.LineStyle = xlContinuous
                If InsZ Like "*Med*" Then borderRange.Borders.Weight = xlMedium
                If InsZ Like "*Thick*" Then borderRange.Borders(xlEdgeLeft).Weight = xlThick
                If InsZ Like "*Double*" Then borderRange.Borders.LineStyle = xlDouble
            End If
            If InsZ = "Banding" Then
                Dim xx As Integer, xColor As Integer
                xColor = 0
                For xx = dataStartRow To dataEndRow + 1
                    If xColor = 255 Then xColor = 230 Else xColor = 255
                    xlSheet.Range("A" & xx & ":" & Num2ColumnLetter(ColumnCount) & xx).Interior.Color = RGB(xColor, xColor, xColor)
                Next xx
            End If
            If InsZ Like "*cBorder*" Then
                Dim BorderColor As Long, BorderWeight As Long
                Dim Endcol As String
                BorderColor = RGB(0, 0, 0)
                BorderWeight = xlMedium
                If InsZ Like "*Thin*" Then BorderWeight = xlThin
                If InsZ Like "*Med*" Then BorderWeight = xlMedium
                If InsZ Like "*Thick*" Then BorderWeight = xlThick
                If InsZ Like "*RedBorder*" Then BorderColor = RGB(255, 0, 0)
                Endcol = IIf(InsZ Like "*End*", SplitString(SplitString(InsZ, "End=", 1), ":", 0), Colz)
                Dim TargetRange As Object
                Set TargetRange = xlSheet.Range(Colz & (dataStartRow - 1) & ":" & Endcol & dataEndRow)
                With TargetRange
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).Weight = BorderWeight
                    .Borders(xlEdgeRight).Weight = BorderWeight
                    .Borders(xlEdgeTop).Weight = BorderWeight
                    .Borders(xlEdgeBottom).Weight = BorderWeight
                    .Borders(xlEdgeLeft).Color = BorderColor
                    .Borders(xlEdgeRight).Color = BorderColor
                    .Borders(xlEdgeTop).Color = BorderColor
                    .Borders(xlEdgeBottom).Color = BorderColor
                End With
            End If
            If InsZ Like "*SubTotal*" Then
                Dim dRng As Object, StRow As Long, EndRow As Long
                Dim TempStr As String
                Dim TotColumns(1 To 4) As Long
                EndColz = SplitString(SplitString(InsZ, "SubTotal", 1), ":", 0)
                TempStr = SplitString(InsZ, ":", 1)
                TempStr = Replace(TempStr, "-", ", ")
                StRow = dataStartRow
                EndRow = xlSheet.Range(Colz & xlSheet.rows.Count).End(xlUp).Row
                Set dRng = xlSheet.Range(Colz & StRow & ":" & EndColz & EndRow)
                dRng.Sort Key1:=xlSheet.Range(Colz & dataStartRow + 1), Order1:=xlAscending, Header:=xlYes
                For i = 1 To 4
                    TotColumns(i) = i
                Next i
                dRng.SubTotal GroupBY:=1, Function:=xlSum, TotalList:=TotColumns, Replace:=True
            End If
            If InsZ Like "FxAll*" Then
                FormX = Split(InsZ, "FxAll")(1)
                xlApp.Calculation = xlManual
                dataRange.formula = FormX
                xlApp.Calculation = xlCalculationAutomatic
            End If
            If InsZ Like "*cFormula*" Then
                Dim FormulaZ As String, fFormat As String
                EndColz = SplitString(SplitString(InsZ, "cFormula", 1), ":", 0)
                RowZ = SplitString(SplitString(InsZ, "cFormula", 1), ":", 1)
                FormulaZ = SplitString(SplitString(InsZ, "cFormula", 1), ":", 2)
                fFormat = SplitString(SplitString(InsZ, "cFormula", 1), ":", 3)
                If EndColz & "test" = "test" Then EndColz = Colz
                Dim formulaRange As Range
                Set formulaRange = xlSheet.Range(Colz & RowZ & ":" & EndColz & RowZ)
                formulaRange.formula = FormulaZ
                If fFormat Like "$*" Then formulaRange.Style = "Currency"
                If fFormat Like "$Int*" Then formulaRange.NumberFormat = "$#,##0_);[Red]($#,##0)"
                If fFormat Like "S*" Then formulaRange.NumberFormat = "#,##"
                If fFormat Like "SInt*" Then formulaRange.NumberFormat = "#,##0_);[Red](#,##0)"
                If fFormat Like "%*" Then formulaRange.Style = "Percent"
                If fFormat Like "P1*" Then formulaRange.NumberFormat = "0.0%"
                If fFormat Like "P2*" Then formulaRange.NumberFormat = "0.00%"
                If fFormat Like "P3*" Then formulaRange.NumberFormat = "0.000%"
                If fFormat Like "P4*" Then formulaRange.NumberFormat = "0.0000%"
                If fFormat Like "P5*" Then formulaRange.NumberFormat = "0.00000%"
                If fFormat Like "P6*" Then formulaRange.NumberFormat = "0.000000%"
                If fFormat Like "P7*" Then formulaRange.NumberFormat = "0.0000000%"
                If fFormat Like "*B*" Then formulaRange.Font.Bold = True
                If val(RowZ) < dataStartRow Then formulaRange.HorizontalAlignment = xlCenter
            End If
            If InsZ Like "ColFntSize:*" Then dataRange.Font.Size = val(SplitString(InsZ, "ColFntSize:", 1))
            If InsZ Like "*cColor*" Then
                EndColz = IIf(InsZ Like "*End=*", SplitString(SplitString(InsZ, "End=", 1), ":", 0), Colz)
                RowZ = SplitString(SplitString(InsZ, "cColor", 1), ":", 1)
                ColorZ = SplitString(SplitString(InsZ, "cColor", 1), ":", 2)
                If InStr(ColorZ, ".") = 0 Then ColorZ = GetRGBFromName(ColorZ)
                Dim colorRange As Object
                Set colorRange = xlSheet.Range(Colz & RowZ & ":" & EndColz & RowZ)
                colorRange.Interior.Color = RGB(val(SplitString(ColorZ, ".", 0)), val(SplitString(ColorZ, ".", 1)), val(SplitString(ColorZ, ".", 2)))
            End If
            If InsZ Like "*cFontColor*" Then
                EndColz = IIf(InsZ Like "*End=*", SplitString(SplitString(InsZ, "End=", 1), ":", 0), Colz)
                RowZ = SplitString(SplitString(InsZ, "cFontColor", 1), ":", 1)
                ColorZ = SplitString(SplitString(InsZ, "cFontColor", 1), ":", 2)
                If InStr(ColorZ, ".") = 0 Then ColorZ = GetRGBFromName(ColorZ)
                Dim fontColorRange As Object
                Set fontColorRange = xlSheet.Range(Colz & RowZ & ":" & EndColz & RowZ)
                fontColorRange.Font.Color = RGB(val(SplitString(ColorZ, ".", 0)), val(SplitString(ColorZ, ".", 1)), val(SplitString(ColorZ, ".", 2)))
            End If
            If InsZ Like "*VColor*" Then
                EndColz = SplitString(SplitString(InsZ, "Color", 1), ":", 0)
                RowZ = SplitString(SplitString(InsZ, "Color", 1), ":", 1)
                ColorZ = SplitString(SplitString(InsZ, "Color", 1), ":", 2)
                If InStr(ColorZ, ".") = 0 Then ColorZ = GetRGBFromName(ColorZ)
                xlSheet.Range(Colz & dataStartRow & ":" & Colz & RowZ).Interior.Color = RGB(val(SplitString(ColorZ, ".", 0)), val(SplitString(ColorZ, ".", 1)), val(SplitString(ColorZ, ".", 2)))
            End If
            If InsZ Like "*FColor*" Then
                ColorZ = SplitString(InsZ, ":", 1)
                If InStr(ColorZ, ".") = 0 Then ColorZ = GetRGBFromName(ColorZ)
                dataRange.Font.Color = RGB(val(SplitString(ColorZ, ".", 0)), val(SplitString(ColorZ, ".", 1)), val(SplitString(ColorZ, ".", 2)))
            End If
            If InsZ Like "*InsertRow*" Then
                EndColz = SplitString(SplitString(InsZ, "InsertRow", 1), ":", 0)
                RowZ = SplitString(SplitString(InsZ, "InsertRow", 1), ":", 1)
                xlSheet.Range(Colz & RowZ & ":" & Colz & RowZ).EntireRow.Insert
            End If
            ' Replaced with 'cLabel' 2/26/2026 Leo Johnson
            'If InsZ Like "Merge*" Then
            '    EndColz = IIf(InsZ Like "*End*", SplitString(SplitString(InsZ, "End=", 1), ":", 0), Colz)
            '    EndColz = SplitString(SplitString(InsZ, "Merge", 1), ":", 0)
            '    RowZ = SplitString(SplitString(InsZ, "Merge", 1), ":", 1)
            '    Textz = SplitString(SplitString(InsZ, "Merge", 1), ":", 2)
            '
            '   ' Dim mergeRange As Object
            '    Set mergeRange = xlSheet.Range(Colz & RowZ & ":" & EndColz & RowZ)
            '    If EndColz <> Colz Then mergeRange.Merge
            '    mergeRange.VerticalAlignment = xlCenter
            '    mergeRange.Value = Textz
            'End If
            If InsZ Like "cLabel*" Then
                If InsZ Like "*End*" Then
                EndColz = SplitString(SplitString(InsZ, "End=", 1), ":", 0)
                   RowZ = SplitString(SplitString(InsZ, "End=", 1), ":", 1)
                  Textz = SplitString(SplitString(InsZ, "End=", 1), ":", 2)
                Else
                  EndColz = Colz
                  RowZ = SplitString(SplitString(InsZ, "cLabel", 1), ":", 1)
                  Textz = SplitString(SplitString(InsZ, "cLabel", 1), ":", 2)
                End If
                Dim mergeRange As Object
                Set mergeRange = xlSheet.Range(Colz & RowZ & ":" & EndColz & RowZ)
                If EndColz <> Colz Then mergeRange.Merge
                mergeRange.VerticalAlignment = xlCenter
                mergeRange.Value = Textz
            End If
            If InsZ <> "WRAP" And InsZ Like "W*" Then dataRange.ColumnWidth = val(SplitString(InsZ, "W", 1))
            If InsZ Like "DisplayAll*" Then xlSheet.Columns.AutoFit
            If InsZ Like "TextDateOff*" Then xlApp.ErrorCheckingOptions.TextDate = False
            If InsZ Like "NumberAsTextOff*" Then xlApp.ErrorCheckingOptions.NumberAsText = False
            If InsZ Like "BackgroundCheckingOff*" Then xlApp.ErrorCheckingOptions.BackgroundChecking = False
            If InsZ Like "CondFormat*" Then
                Dim cfRange As Object
                Dim cfColorR As Long, cfColorG As Long, cfColorB As Long
                Dim fontColorR As Long, fontColorG As Long, fontColorB As Long
                Dim isBold As Boolean
                Dim condParts() As String, ii As Long
                Set cfRange = xlSheet.Range(Colz & dataStartRow & ":" & Colz & dataEndRow)
                If InsZ Like "*Color.RGB*" Then
                    Dim rgbStr As String
                    rgbStr = SplitString(SplitString(InsZ, "Color.RGB(", 1), ")", 0)
                    cfColorR = val(SplitString(rgbStr, ".", 0))
                    cfColorG = val(SplitString(rgbStr, ".", 1))
                    cfColorB = val(SplitString(rgbStr, ".", 2))
                End If
                If InsZ Like "*FontColor.RGB*" Then
                    Dim fStr As String
                    fStr = SplitString(SplitString(InsZ, "FontColor.RGB(", 1), ")", 0)
                    fontColorR = val(SplitString(fStr, ".", 0))
                    fontColorG = val(SplitString(fStr, ".", 1))
                    fontColorB = val(SplitString(fStr, ".", 2))
                End If
                isBold = (InStr(InsZ, "Font.Bold=True") > 0)
                condParts = Split(InsZ, "Range")
                For ii = 1 To UBound(condParts)
                    Dim condStr As String, op As String, valx As Double
                    condStr = Split(condParts(ii), ":")(0)
                    If InStr(condStr, ">") > 0 Then
                        op = ">"
                        valx = val(Split(condStr, ">")(1))
                        With cfRange.FormatConditions.Add(Type:=1, Operator:=5, Formula1:=valx)
                            .Interior.Color = RGB(cfColorR, cfColorG, cfColorB)
                            If fontColorR + fontColorG + fontColorB > 0 Then .Font.Color = RGB(fontColorR, fontColorG, fontColorB)
                            If isBold Then .Font.Bold = True
                        End With
                    ElseIf InStr(condStr, "<") > 0 Then
                        op = "<"
                        valx = val(Split(condStr, "<")(1))
                        With cfRange.FormatConditions.Add(Type:=1, Operator:=1, Formula1:=valx)
                            .Interior.Color = RGB(cfColorR, cfColorG, cfColorB)
                            If fontColorR + fontColorG + fontColorB > 0 Then .Font.Color = RGB(fontColorR, fontColorG, fontColorB)
                            If isBold Then .Font.Bold = True
                        End With
                    ElseIf InStr(condStr, ".") > 0 Then
                        Dim bounds() As String
                        bounds = Split(condStr, ".")
                        If UBound(bounds) = 1 Then
                            Dim val1 As Double, val2 As Double
                            val1 = val(bounds(0))
                            val2 = val(bounds(1))
                            With cfRange.FormatConditions.Add(Type:=1, Operator:=3, Formula1:=val1, Formula2:=val2)
                                .Interior.Color = RGB(cfColorR, cfColorG, cfColorB)
                                If fontColorR + fontColorG + fontColorB > 0 Then .Font.Color = RGB(fontColorR, fontColorG, fontColorB)
                                If isBold Then .Font.Bold = True
                            End With
                        End If
                    End If
                Next ii
            End If
        Next LoopX
    End If
    If IsOpen("MessagePopUp") = True Then MessageUser "Update", "Posting: " & sTabName, "Setting Page Parameters..."
    With xlSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = Left(WorkSheetTitleX, 100)
        .CenterFooter = "© " & Year(Date) & " Fleet Advantage - Internal Use Only"
        .LeftMargin = xlApp.InchesToPoints(0.25)
        .RightMargin = xlApp.InchesToPoints(0.25)
        .TopMargin = xlApp.InchesToPoints(0.166)
        .BottomMargin = xlApp.InchesToPoints(0.35)
        .HeaderMargin = xlApp.InchesToPoints(0.166)
        .FooterMargin = xlApp.InchesToPoints(0#)
        .PrintTitleRows = "$1:$3"
        .PrintHeadings = False
        .PrintGridlines = False
    End With
    xlSheet.Range("B1").ColumnWidth = 15
    If GroupReports = 0 Or LastReport = True Then xlApp.Visible = True
    If SaveMe = True Then
        xlApp.DisplayAlerts = False
        If Not sFile Like "*XLSX*" Then sFile = "C:\Faas\" & FaUserName & "\Temp\" & sFile & ".xlsx"
        xlWBk.SaveAs fileName:=sFile
        xlApp.DisplayAlerts = True
    End If
    If LastReport <> True And GroupReports <> 0 Then
        xlApp.DisplayAlerts = False
        xlWBk.SaveAs fileName:=sFile
        xlApp.DisplayAlerts = True
        xlWBk.close
        xlApp.Quit
    End If
    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True
    Exit Function
Err_Handler:
    MsgBox Err.Description, vbExclamation, Err.Number
    If Not rs Is Nothing Then rs.close
    Set rs = Nothing
    DoCmd.SetWarnings True
    Exit Function
End Function
Function GetRGBFromName(ColorName As String) As Variant
Select Case LCase(Trim(ColorName))
    Case "red": GetRGBFromName = "255.0.0"
    Case "green": GetRGBFromName = "0.128.0"
    Case "blue": GetRGBFromName = "0.0.255"
    Case "yellow": GetRGBFromName = "255.255.0"
    Case "black": GetRGBFromName = "0.0.0"
    Case "white": GetRGBFromName = "255.255.255"
    Case "gray", "grey": GetRGBFromName = "128.128.128"
    Case "light gray", "light grey": GetRGBFromName = "211.211.211"
    Case "dark gray", "dark grey": GetRGBFromName = "64.64.64"
    Case "orange": GetRGBFromName = "255.165.0"
    Case "purple": GetRGBFromName = "128.0.128"
    Case "pink": GetRGBFromName = "255.192.203"
    Case "brown": GetRGBFromName = "165.42.42"
    Case "cyan": GetRGBFromName = "0.255.255"
    Case "magenta": GetRGBFromName = "255.0.255"
    Case "lime": GetRGBFromName = "0.255.0"
    Case "navy": GetRGBFromName = "0.0.128"
    Case "teal": GetRGBFromName = "0.128.128"
    Case "olive": GetRGBFromName = "128.128.0"
    Case "maroon": GetRGBFromName = "128.0.0"
    Case "gold": GetRGBFromName = "255.215.0"
    Case "silver": GetRGBFromName = "192.192.192"
    Case "beige": GetRGBFromName = "245.245.220"
    Case "indigo": GetRGBFromName = "75.0.130"
    Case "violet": GetRGBFromName = "238.130.238"
    Case "turquoise": GetRGBFromName = "64.224.208"
    Case "salmon": GetRGBFromName = "250.128.114"
    Case "coral": GetRGBFromName = "255.127.80"
    Case "chocolate": GetRGBFromName = "210.105.30"
    Case "crimson": GetRGBFromName = "220.20.60"
    Case "light green": GetRGBFromName = "144.238.144"
    Case "dark green": GetRGBFromName = "0.100.0"
    Case "light blue": GetRGBFromName = "173.216.230"
    Case "dark blue": GetRGBFromName = "0.0.139"
    Case "sky blue": GetRGBFromName = "135.206.235"
    Case "peach": GetRGBFromName = "255.218.185"
    Case "lavender": GetRGBFromName = "230.230.250"
    Case "tan": GetRGBFromName = "210.180.140"
    Case "khaki": GetRGBFromName = "240.230.140"
    Case "aquamarine": GetRGBFromName = "127.255.212"
    Case "azure": GetRGBFromName = "240.255.255"
    Case "mint": GetRGBFromName = "189.252.201"
    Case Else: GetRGBFromName = "255.255.255" ' Default to white
End Select

End Function

Function Num2ColumnLetter(ColumnNumber As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    Num2ColumnLetter = s
End Function
Function MakeDocsReport(Optional formName As String)
    Dim TitleX, XLS As String
    Dim Countx As Integer
    Countx = DCount("AssetId", "ScheduleUnitDocs4")
    TitleX = ""
    TitleX = TitleX & DLookup("[Lease No]", "ScheduleUnitDocs4") & " | "
    TitleX = TitleX & "Sch: " & DLookup("SchNo", "ScheduleUnitDocs4") & " | "
    TitleX = TitleX & "Units: " & Countx & " "
    XLS = "|D,BC|E,BC|F,TC|I,$E|J,TC|L,CX|M,TC|N,$E|R,$E|V,$E|Z,$E" & _
                    "|AD,$E|AH,$E|AL,$E|AP,$E|AT,$E|AX,$E" & _
                    "|BB,$E|BF,$E|BJ,$E|BN,$E|BR,$E|BV,$E|BZ,$E" & _
                    "|CD,$E|CH,$E|CL,$E|CP,$E|CS,$E|CU,TC|CY,BC|DG,$E|"
    FormatExcel "ScheduleUnitDocs4", "C:\Faas\" & FaUserName & "\Temp\ScheduleUnitDocs4.xlsx", "Unit_Docs", , "" & XLS & "", Countx, "" & TitleX & "", True
    'XLS = Col letter,Format Option
    ' first set = number of columns
    'Format Options Type Command
    'Types:     T = Text | B = Bold | $ = Currency | P = Percent
    'Commands:  C = Center | L=Lexf | R=Right | E = Sum col | T= Total Col Count | X = COUNT | W10 set column width
    '|[Column],[Type][Command]|
End Function
Function LastTitleEventReport(iSchID As Long)
    SQLz = "Select S.ClientShortName As Client"
    SQLz = SQLz & ", S.MLNo"
    SQLz = SQLz & ", S.Schedule"
    SQLz = SQLz & ", S.UnitGroup AS [Group]"
    SQLz = SQLz & ", S.AssetID"
    SQLz = SQLz & ", S.VIN"
    SQLz = SQLz & ", UnitTitleLogDate as LastLog"
    SQLz = SQLz & ", TitleNumber as Title"
    SQLz = SQLz & ", EntryReason as [Action]"
    SQLz = SQLz & ", TitledLocationState as TitleState "
    SQLz = SQLz & "FROM (vw_SixKeys AS S "
    SQLz = SQLz & "LEFT OUTER JOIN UnitTitleHistoryLastEntry As H on S.AssetID = H.AssetID) "
    SQLz = SQLz & "WHERE S.SchID = " & iSchID & " "
    SQLz = SQLz & "ORDER BY S.UnitGroup,S.AssetID;"
    '
    '   SQLz = "Select Client, MLNo, Schedule, [Group], AssetID, VIN, UnitTitleLogDate as LastLog,  TitleNumber as Title, EntryReason as [Action],TitledLocationState as TitleState"
    '   SQLz = SQLz & " From UnitTitleHistoryLastEntry"
    '   SQLz = SQLz & " Where SchID = " & iSchId & " "
    '   SQLz = SQLz & " Order By [Group],AssetID;"
    Dim TitleX, XLS As String
    TitleX = DLookup("Client", "vw_FiveKeys", "SchID=" & iSchID) & ", Sch: " & DLookup("Schedule", "vw_FiveKeys", "SchID=" & iSchID) & "         Last Title Action " & Date
    XLS = "|E,BC|J,BC|A,LL2|A,W15|"
    SendReport2Excel SQLz, "Last Action", TitleX, True, XLS

End Function

Function MakeOPSReport(Optional ClientNamex As String, Optional SchIdi As Long, Optional ShortReport As Boolean)
    Dim XLS As String
    Dim TitleX As String
    Dim ScheduleX As Long
    Dim RST    As DAO.Recordset
    MessageUser "Open", "Ops Report", "This will take just a moment...", False
    If SchIdi = 0 Then
        ScheduleX = DMax("ScheduleID", "Schedule_Group_WorkTable")
    Else
        ScheduleX = SchIdi
    End If
    SQLz = "Select Client, MLNo, Schedule From vw_FiveKeys Where SchId=" & ScheduleX & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    TitleX = "OPS Sheet - " & Nz(RST!Client, "") & " - " & Nz(RST!MLNo, "") & " Sch: " & Nz(RST!Schedule, "")
    RST.close
    If ShortReport = True Then
        SQLz = "Select Schedule"                                    'Col A
        SQLz = SQLz & ", UnitGroup As[Group]"                       'Col B
        SQLz = SQLz & ", EsitmatedDeliveryDate AS [Est Delv]"       'Col C
        SQLz = SQLz & ", ActualDeliveryDate as DelvDate"            'Col D
        SQLz = SQLz & ", UnitDescShortVer as Description"           'Col E
        SQLz = SQLz & ", VIN"                                       'Col F
        SQLz = SQLz & ", UnitNum as [Unit Num]"                     'Col G
        SQLz = SQLz & ", R.LocationName as [Reg Location]"          'Col H
        SQLz = SQLz & ", L.LocationName as Domicile"                'Col I
        SQLz = SQLz & ", L.LAddrLine1 as StreetAddress"             'Col J
        SQLz = SQLz & ", L.LCity as City"                           'Col K
        SQLz = SQLz & ", L.LState as State"                         'Col L
        SQLz = SQLz & ", L.LZip as Zip"                             'Col M
        SQLz = SQLz & ", L.LCounty AS County "                      'Col N
        SQLz = SQLz & ", '' As [BasePlate?] "                       'Col O
        SQLz = SQLz & ", UnitCost"                                  'Col P
        SQLz = SQLz & ", PartsCost"                                 'Col Q
        SQLz = SQLz & ", U.VendorCost"                              'Col R
        SQLz = SQLz & ", U.UnitPlateNum As PlateNum"                'Col S
        SQLz = SQLz & ", UnitIRP"                                   'Col T
        SQLz = SQLz & ", AssetID  "                                 'Col U
        SQLz = SQLz & ", IIF(U.UnitTaxExempt = 0,'Rent Stream','Exempt')                     AS [Tax Due]"          'Col V
        SQLz = SQLz & ", Format(LTax.StateSales,        'Percent')                           AS [State Tax]"        'Col W
        SQLz = SQLz & ", Format(LTax.CountySales,       'Percent')                           AS [County Tax]"       'Col X
        SQLz = SQLz & ", Format(LTax.CitySales,         'Percent')                           AS [City Tax]"         'Col Y
        SQLz = SQLz & ", Format( ROUND(L.LTaxRate - NZ(LTax.LSalesTax, 0), 3), 'Percent')    AS [Other Tax]"        'Col Z
        SQLz = SQLz & ", Format(L.LTaxRate,             'Percent')                           AS [Loc Tax]"          'Col AA
        SQLz = SQLz & ", U.UnitTaxNotes  AS [Tax Notes]"                                                            'Col AB
        SQLz = SQLz & ", L.LastTaxUpdate  AS [Last Update]"                                                         'Col AC
        SQLz = SQLz & " FROM ((((vw_SixKeys AS S "
        SQLz = SQLz & " INNER JOIN Units AS U ON S.UnitID = U.UnitID) "
        SQLz = SQLz & " LEFT OUTER JOIN Locations AS L ON S.LocationID = L.LocationID) "
        SQLz = SQLz & " LEFT OUTER JOIN Locations AS R ON U.UnitRegLocationRef = R.LocationID) "
        SQLz = SQLz & " LEFT OUTER JOIN (SELECT LocationID, StateSales, CountySales, CitySales, LSalesTax"
        SQLz = SQLz & " FROM vw_LocationTaxSales2) AS LTax ON L.LocationID = LTax.LocationID)"
        SQLz = SQLz & " WHERE SchID = " & ScheduleX & " Order By UnitNum, UnitGroup"
        TitleX = TitleX & " Short Report"
        XLS = "|A,TC|B,TC|C,TC|D,TC|L,BC|P,BC|P,$E|Q,$E|R,$E|T,TX|U,BC|V,TC|W,P3|X,P3|Y,P3|Z,P3|AA,P3"
        XLS = XLS & "|W,TC|X,TC|Y,TC|Z,TC|AA,TC|V,cBorder:Med:End=AB|Z,cBorder:Med|Z,cColor:3:crimson|AA,cColor:3:black|AC,cBorder:Med|"
    Else         ' Long Report
        SQLz = "Select * From OPS_Sheet Where SchID = " & ScheduleX & " Order By AssetId"
        XLS = "|B,TC|D,TC|F,TX|G,TC|J,TL|M,TC|Z,TC|AD,$E|AE,$E|AF,$E|AG,$E|AH,$E|AI,$E|AJ,$E|AQ,SR|AR,TC|" ''AV,HH|AV,HH|"
        If DCount("AssetID", "Cost_ChildParts", "SchID=" & ScheduleX & " And PartClass Like '*Temp*'") < 1 Then XLS = XLS & "Q,ZZ|Q,ZZ|Q,ZZ|"
    End If
    Set RST = Nothing
    SendReport2Excel SQLz, "OPS Report", TitleX, True, XLS, , , , , , "C:\FAAS\" & FaUserName & "\Temp\OPSReport.xlsx"
    MessageUser ("Close")
End Function


Function MakeScheduleUnitSummary(ClientIDx As Integer, SchIDx As Integer, ScheduleX As String)
    Dim XLS, ClientX As String
    Dim Countx As Integer
    ClientX = DLookup("clientCompanyName", "clients", "clientID=" & ClientIDx)
    SQLz = "Select * From ScheduleUnitSummary Where Schid=" & SchIDx & ";"
    XLS = "|A,DD|A,BC|B,BC|H,$E|I,$E|K,BC|M,BC|M,W15|N,W15|O,$E|O,W20|P,BC|P,W10" & _
                    "|Q,$E|Q,W20|R,BC|R,W10|S,$E|S,W15|T,$E|T,W15|U,HH|V,HH|W,HH|"
    SendReport2Excel SQLz, "Summary", "Schedule-Unit Summary: " & ClientX & " Schedule " & ScheduleX & " Units: " & Countx, True, XLS, Countx, , , , , "C:\FAAS\" & FaUserName & "\Temp\ScheduleUnitSummary.xlsx"
End Function
Public Function IfExcelOpenCloseIt(FileNameX As String)
    Dim oApp   As Object
    Dim xl     As Integer
    Dim FileOpen As Boolean
    Set oApp = GetApplication("excel.Application")
    If oApp Is Nothing Then Exit Function
    FileOpen = False
    For xl = 1 To oApp.Workbooks.Count
        If oApp.Workbooks(xl).FullName = FileNameX Then FileOpen = True
        If FileOpen = True Then
            oApp.Workbooks(xl).Save
            oApp.Workbooks(xl).close
            Exit For
        End If
    Next xl
    Set oApp = Nothing
End Function

Private Function GetApplication(ByVal AppClass As String) As Object
    Const vbErr_AppNotRun = 429
    On Error Resume Next
    Set GetApplication = GetObject(Class:=AppClass)
    If Err.Number = vbErr_AppNotRun Then Set GetApplication = CreateObject(AppClass)
    On Error GoTo 0
End Function
Public Function ExportFleetList(NameToUse As String, ClientGroupList As Boolean, ClientIDx As Long, ClientGroupIDx As Long)
    Dim CustomScript, Header, WhereX, TabX As String
    Dim x, UnitsX As Integer
    Dim RST    As DAO.Recordset
    'On Error GoTo Fixme_Error
     On Error GoTo 0
    CustomScript = "|B,BC|F,CX|I,BC|L,BL|O,BC|Q,$E|R,$E|S,$E|T,$E|V,$R|W,P4C|AC,LX|AH,SR|AL,SR|AN,SR|AQ,SR|AR,$R|AS,SR|AT,$R|"
    Call MessageUser("Open", "Creating Report", " ...This will take just a moment...")

    If ClientGroupList = True Then
        Header = "Client Group (" & NameToUse & ")- Fleet List-"
        TabX = "Grouped Fleet List"
        WhereX = "WHERE  ClientGroupID=" & ClientGroupIDx & " "
    Else
        Header = "Client " & NameToUse & " Fleet List-"
        TabX = "Client Fleet List"
        WhereX = "WHERE  ClientID=" & ClientIDx & " "
    End If

    SQLz = "SELECT U.UnitStatus, [UnitStatusDesc], Units From LookUp_UnitStatus as L "
    SQLz = SQLz & "Inner Join (SELECT UnitStatus,Count(UnitId) As Units "
    SQLz = SQLz & "FROM Units "
    SQLz = SQLz & WhereX
    SQLz = SQLz & "GROUP BY UnitStatus) As U On L.UnitStatus=U.unitStatus "
    SQLz = SQLz & "ORDER BY U.UnitStatus"

    Call MessageUser("Update", "Done doing Unit counts...", " ...Just a Bit More...")
    Set RST = CurrentDb.OpenRecordset(SQLz)
    RST.MoveLast
    RST.MoveFirst
    For x = 1 To RST.RecordCount
        Header = Header & RST!UnitStatusDesc & ":" & RST!UNITS & " "
        UnitsX = UnitsX + RST!UNITS
        RST.MoveNext
    Next x
    Header = Header & "As of:" & Date
    RST.close
    Set RST = Nothing
    Call MessageUser("Update", "Sending data to Excel...", " ...Excel will open soon...")
    SQLz = "SELECT F.Client"                 'A
    SQLz = SQLz & ", F.LeaseType AS Lease"
    SQLz = SQLz & ", F.MLNo"
    SQLz = SQLz & ", F.Schedule"             'D
    SQLz = SQLz & ", F.UnitGroup"
    SQLz = SQLz & ", F.AssetID"
    SQLz = SQLz & ", F.[Unit Number]"
    SQLz = SQLz & ", F.VIN"                  'H
    SQLz = SQLz & ", F.MYear"
    SQLz = SQLz & ", F.Make"
    SQLz = SQLz & ", F.Model"
    SQLz = SQLz & ", F.Type"                 'L
    SQLz = SQLz & ", F.[Sub Type]"
    SQLz = SQLz & ", F.Axle"
    SQLz = SQLz & ", F.Status"
    SQLz = SQLz & ", F.[In Service]"         'P
    SQLz = SQLz & ", F.[Unit Cost]"
    SQLz = SQLz & ", F.[ChildPart Cost]"
    SQLz = SQLz & ", F.[Total OEC]"
    SQLz = SQLz & ", F.CurrRent"             'T
    SQLz = SQLz & ", F.CurrRentType"
    SQLz = SQLz & ", F.[Synd Residual]"
    SQLz = SQLz & ", F.LRF"
    SQLz = SQLz & ", F.[Service Fee]"
    SQLz = SQLz & ", F.[Lease Start]"
    SQLz = SQLz & ", F.[Lease Term]"         'Z
    SQLz = SQLz & ", F.[Lease End]"          'AA
    SQLz = SQLz & ", F.Assignee"
    SQLz = SQLz & ", F.DC_Garage"
    SQLz = SQLz & ", F.DivisionName"         'AD
    SQLz = SQLz & ", F.PlateNum"
    SQLz = SQLz & ", F.PlateExpire"
    SQLz = SQLz & ", F.LTaxRate"             'AG
    SQLz = SQLz & ", F.LTDFuel"              'AH
    SQLz = SQLz & ", F.LTDMPG"
    SQLz = SQLz & ", F.LTDFuelMov"
    SQLz = SQLz & ", F.LTDMPGMov"
    SQLz = SQLz & ", F.OldestReport"
    SQLz = SQLz & ", F.LastReport"           'AN
    SQLz = SQLz & ", F.OBCMiles"
    SQLz = SQLz & ", F.ODORead AS Odometer"
    SQLz = SQLz & ", F.[1stExchangeITDate]"
    SQLz = SQLz & ", F.MilesChrgAfterT1"     'AR
    SQLz = SQLz & ", F.ChrgedAmtT1"
    SQLz = SQLz & ", F.MilesChrgAfterT2"
    SQLz = SQLz & ", F.ChrgedAmtT2"
    SQLz = SQLz & ", F.DelvLoc"
    SQLz = SQLz & ", F.clientID"
    SQLz = SQLz & ", F.ClientGroupId"
    SQLz = SQLz & ", F.Engine "
    SQLz = SQLz & "FROM vw_ClientFleetList As F " & WhereX & " "
    SQLz = SQLz & "ORDER BY Client, LeaseType, Schedule, UnitGroup "
    'UnitCount = CurrentDb.OpenRecordset(SQLz).RecordCount

    SendReport2Excel SQLz, TabX, Header, True, CustomScript, 0, False

    Call MessageUser("Close")

    On Error GoTo 0
    Exit Function
Fixme_Error:
    Call MessageUser("Close")
    MsgBox "Sorry, But " & NameToUse & " dosen't have a Fleet List", vbCritical, "No Units"
End Function
