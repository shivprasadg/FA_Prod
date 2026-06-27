Attribute VB_Name = "SQLServerCalls"
Option Compare Database
Option Explicit

Public Function cConnection(element As Integer) As String
'element def >>> 0 = Server 1=UIDType 2 = Database
    Dim TestConn As String
    Dim Server, UIDType As String
    TestConn = CurrentDb.TableDefs("units").Connect
    UIDType = "Trusted_Connection=yes"
    If TestConn Like "*Phoenix*" Then Server = "FA-Phoenix; "
    If TestConn Like "*10.1.0.7*" Then Server = "10.1.0.7,1549; "
    If TestConn Like "*fa-poseidon.domain.local*" Then Server = "fa-poseidon.domain.local,1549; "
    If TestConn Like "*Poseidon.eastus.*" Then Server = "fa-poseidon.eastus.cloudapp.azure.com,1549; "
    If TestConn Like "*Poseidon.eastus.*" Then UIDType = "UID=" & Decrypt("vtyqx}yx") & ";Pwd=" & Decrypt("_|‡pyx€v…A=") & ""
    If element = 0 Then cConnection = Server
    If element = 1 Then cConnection = UIDType
    If element = 2 Then cConnection = SplitString(SplitString(TestConn, "Database=", 1), ";", 0)
    If element = 3 Then cConnection = "MARS Connection=True"
    If element = 4 Then cConnection = "16384"
End Function

Public Function CloneScheduleGroup(SchIdi As Integer, GroupIDi As Integer, SchNoX As String _
                                  , GrpNoX As String, SchDescX As String, GrpDescX As String _
                                  , Optional JustGroupi As Integer, Optional NewSchIDi As Integer, Optional JustSchi As Integer)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "CloneScheduleGroup "

        .Parameters.Append .CreateParameter("@SourceSchID", adInteger, adParamInput, 3, SchIdi)
        .Parameters.Append .CreateParameter("@SourceGroupID", adInteger, adParamInput, 3, GroupIDi)
        .Parameters.Append .CreateParameter("@NewSchName", adVarChar, adParamInput, 100, SchNoX)
        .Parameters.Append .CreateParameter("@NewSchDesc", adVarChar, adParamInput, 100, SchDescX)
        .Parameters.Append .CreateParameter("@NewGrpName", adVarChar, adParamInput, 100, GrpNoX)
        .Parameters.Append .CreateParameter("@NewGrpDesc", adVarChar, adParamInput, 100, GrpDescX)
        .Parameters.Append .CreateParameter("@JustGroup", adInteger, adParamInput, 3, JustGroupi)
        .Parameters.Append .CreateParameter("@NewSchID", adInteger, adParamInput, 3, NewSchIDi)
        .Parameters.Append .CreateParameter("@JustSch", adInteger, adParamInput, 3, JustSchi)
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
End Function

Public Function RePairExchangeIT()
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "RePairExchangeIT"
        ' .Parameters.Append .CreateParameter("@AssetsToUpdate", adVarChar, adParamInput, 255, AssetsToUpdate)
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
End Function
Public Sub SyncUnitsToLocationDivisions()
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "sp_SyncUnitsToLocationDivisions"
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
End Sub

Public Function SyncAuditTables()
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
    '' .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "Quick_Sync_TableAudit"
        ' .Parameters.Append .CreateParameter("@AssetsToUpdate", adVarChar, adParamInput, 255, AssetsToUpdate)
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
End Function

Public Function RmktGainLossFix(noMsg As Boolean)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & "; TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "RmktGainLossFix"
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
    If noMsg = False Then MsgBox "Auto RMKT Gain Loss Proc Ran.", vbOKOnly, "ATLAAS Core"
End Function
Public Function AutoUpdateEOL_Statuses(noMsg As Boolean)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & "; TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "AutoUpdateEOL_Statuses"
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
    If noMsg = False Then MsgBox "Auto Update EOL Statuses Proc Ran.", vbOKOnly, "ATLAAS Core"
End Function

Public Function AutoExtendMTM(noMsg As Boolean)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & "; TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "AutoExtendMTM"
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
    If noMsg = False Then MsgBox "Auto Extend Proc Ran.", vbOKOnly, "ATLAAS Core"
End Function
Public Function SetEngineType(Optional noMsg As Boolean)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & "; TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "sp_SetEngineType"
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
    If noMsg = False Then MsgBox "VIN Decode and Engine MFG Fix Done", vbOKOnly, "ATLAAS Core"
End Function
Public Function CloneDecodedVINS() As Integer
    Dim cmd As ADODB.command
    Dim db As String
    Dim iNum As Integer
    Dim pName As String
    Dim RecordsAffected As Long
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & "; TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("Results", adInteger, adParamOutput)
        .CommandText = "sp_CloneDecodedVINS"
        .Execute
    End With
    pName = cmd.Parameters(0).Name
    iNum = cmd.Parameters(0).Value
    cmd.ActiveConnection.close
    CloneDecodedVINS = iNum
    Set cmd = Nothing
   ' If noMsg = False Then MsgBox pName & ": " & iNum & " matched VINs from VINGroup", vbOKOnly, "ATLAAS Core"
End Function
Public Function SetNetBuyOut(Optional noMsg As Boolean)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & "; MARS_Connection=yes; "
        .CommandType = adCmdStoredProc
        .CommandText = "sp_SetNetBuyOut"
        '   .Parameters.Append .CreateParameter("@statement", adVarChar, adParamInput, 255, 1)
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
    If noMsg = False Then MsgBox "Net-BuyOut set", vbOKOnly, "ATLAAS Core"
End Function

Public Function getOPSMetricsReport(Optional startDate As String, Optional Enddate As String) As String
    Dim cmd As ADODB.command
    Dim db As String
    'Dim RST As DAO.Recordset
    Dim RST As Object
    Set cmd = New ADODB.command
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim tbl As ListObject
    Dim rng As Range
    Dim i As Integer

    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "dbo.OPSMetricsReport"
        If startDate <> "" Then .Parameters.Append .CreateParameter("@Start", adDate, adParamInput, 255, CDate(startDate))
        If Enddate <> "" Then .Parameters.Append .CreateParameter("@End", adDate, adParamInput, 255, CDate(Enddate))
        Set RST = .Execute
    End With
    If Not RST.EOF Then
    ' Create a new Excel application
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1) ' Add column headers
            xlApp.Visible = True
        For i = 0 To RST.fields.Count - 1
            xlSheet.Cells(1, i + 1).Value = RST.fields(i).Name
            xlSheet.Cells(2, i + 1).Value = RST.fields(i).Value
        Next i
        xlSheet.Cells.EntireColumn.AutoFit
        Set rng = xlSheet.Range("A1:N2") ' Adjust the range as needed
        Set tbl = xlSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
            tbl.TableStyle = "TableStyleMedium2"
        ' Clean up
        RST.close
        Set RST = Nothing
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        Set tbl = Nothing
        Set rng = Nothing
    cmd.ActiveConnection.close
    Set cmd = Nothing
    Else
        MsgBox "No data found!", vbExclamation
    End If
End Function


Public Function SendServerEmail(ByRef eRecipients As String, _
                                ByRef eCC_recipients As String, _
                                ByRef eSubject As String, _
                                ByRef eBody As String, _
                                Optional eReplyTo As String, _
                                Optional AttachedFiles As String)

    Dim cmd As ADODB.command
    Dim db, eProfile, eFormat As String
    eProfile = "Fleet Advantage"
    eFormat = "HTML"
    Set cmd = New ADODB.command
    db = cConnection(2)

    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandText = "msdb.dbo.sp_send_dbmail @profile_name='" & eProfile & "'" & _
                       ", @recipients = '" & eRecipients & "'" & _
                       ", @copy_recipients = '" & eCC_recipients & "'" & _
                       ", @subject = '" & eSubject & "'" & _
                       ", @body_format = '" & eFormat & "'" & _
                       ", @body= '" & eBody & "'" & _
                       ", @reply_to = '" & eReplyTo & "' "
        If AttachedFiles <> "" Then
            .CommandText = .CommandText & _
                           ", @file_attachments = '" & AttachedFiles & "'"
        End If
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
End Function
Public Function SetExitOptionsLocal(AssetsToUpdate As String)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "sp_SetExitOptionsLocal"
        .Parameters.Append .CreateParameter("@AssetsToUpdate", adVarChar, adParamInput, Len(AssetsToUpdate) + 1, AssetsToUpdate)
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
End Function

Public Function SendRequestPayOffEmail(Optional Emailto As Integer)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "RequestPayOffEmail"
        .CommandTimeout = 300
        .ActiveConnection.CommandTimeout = 300

        If Emailto > 0 Then
            .Parameters.Append .CreateParameter("@EmailTo", adVarChar, adParamInput, 500, FAUserName & "@Fleetadvantage.com")
            If FAUserName = "yjackson" Or FAUserName = "ytj" Then
                .Parameters.Append .CreateParameter("@CopyTo", adVarChar, adParamInput, 500, "ljohnson@fleetadvantage.com;afischer@FleetAdvantage.com")
            Else
                .Parameters.Append .CreateParameter("@CopyTo", adVarChar, adParamInput, 500, "ljohnson@fleetadvantage.com")
            End If
        Else
            .Parameters.Append .CreateParameter("@EmailTo", adVarChar, adParamInput, 500, "yjackson@fleetadvantage.com;afischer@FleetAdvantage.com")
            .Parameters.Append .CreateParameter("@CopyTo", adVarChar, adParamInput, 500, "lmanin@fleetadvantage.com;ljohnson@fleetadvantage.com")
        End If
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
    '  Call ChangeConnectionString(cConnection(0), db, "RemarketingReportAllUnits")
End Function
Public Function SetRemarketingAllUnits()
    On Error GoTo ErrHandler
    Dim cmd As ADODB.command
    Dim db As String
    db = cConnection(2)
    Set cmd = New ADODB.command
    With cmd
        ' Modern OLE DB provider for SQL Server
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "sp_SetRemarketingAllUnits"
        .CommandTimeout = 0
        ' Uncomment if your stored procedure expects parameters
        ' .Parameters.Append .CreateParameter("@AssetsToUpdate", adVarChar, adParamInput, 255, AssetsToUpdate)
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
    ' Update linked Access tables/views
    Call ChangeConnectionString(cConnection(0), db, "RemarketingReportAllUnits")
    Call ChangeConnectionString(cConnection(0), db, "vw_RemarketingReportAllUnits")

    Exit Function
ErrHandler:
    MsgBox "Error in SetRemarketingAllUnits: " & Err.Description, vbCritical
    If Not cmd Is Nothing Then
        If Not cmd.ActiveConnection Is Nothing Then
            cmd.ActiveConnection.close
        End If
        Set cmd = Nothing
    End If
End Function

Public Function FixParkedUnits(noMsg As Boolean)
    Dim cmd As ADODB.command
    Dim db As String
    Set cmd = New ADODB.command
    db = cConnection(2)
    With cmd
        .ActiveConnection = "Provider=sqloledb;Server=" & cConnection(0) & ";Database=" & db & ";" & cConnection(1) & ";ConnectPolicy=Redirect;TrustServerCertificate=yes;MARS_Connection=yes;Network Packet Size=16384;"
        .CommandType = adCmdStoredProc
        .CommandText = "FixParkedLocations"
        .Execute
    End With
    cmd.ActiveConnection.close
    Set cmd = Nothing
    If noMsg = False Then MsgBox "Update Parked Unit Locations Proc Ran.", vbOKOnly, "ATLAAS Core"
End Function
