Attribute VB_Name = "RenameODBConnection"
Option Compare Database
 Public Function ServerTime()
 Set oShell = CreateObject("Shell.Application")
 oShell.ShellExecute "calc.exe", , , "runas", 1
 'ServerTime = Shell("net time")
 End Function

Function DevOrPro() As String
    If WhereAmI Like "*Local*" Then DevOrPro = "Office"
    If WhereAmI Like "*RDS*" Then DevOrPro = "Remote"
    If WhereAmI Like "*OutSide*" Then DevOrPro = "OutSide"
    If WhereAmI Like "*Phoenix*" Then DevOrPro = "Phoenix"
End Function
Function DoesTblExist(strTblName As String) As Boolean
On Error Resume Next
Dim db As DataBase, tbl As TableDef
Set db = CurrentDb
Set tbl = db.TableDefs(strTblName)
If Err.Number = 3265 Then ' Item not found.
DoesTblExist = False
Exit Function
End If
DoesTblExist = True
End Function

Public Function startup()
    Dim lclTBLstr(500) As String
    Dim sqlTBLstr(500) As String
    Dim countER As Integer
  Dim tbl As DAO.TableDef
  Dim tables As DAO.TableDefs
  countER = 1
  Set tables = CurrentDb.TableDefs

 ' Call ChangeConnectionString("fa-poseidon.domain.local,1549", "FleetAdvantage_beSQL", , 0)

'DBEngine.RegisterDatabase "FleetAdvantage_beSQL", "SQL Server", True, "Description=FleetAdvantage_beSQL" & vbCr & "SERVER=fa-poseidon.domain.local,1549" & vbCr & "DATABASE=FleetAdvantage_beSQL"

  For Each tbl In tables
    If tbl.Attributes = dbAttachedODBC Then
        lclTBLstr(countER) = tbl.Name
        sqlTBLstr(countER) = Right(lclTBLstr(countER), Len(lclTBLstr(countER))) ' - 4)
        countER = countER + 1
        tbl.Connect = "Provider=Provider=SQLNCLI11;Data Source=10.1.0.7,1549;Initial Catalog=FleetAdvantage_beSQL;Integrated Security=SSPI;Network Library=DBMSSOCN;"
   '    tbl.RefreshLink
    End If
  Next tbl
  countER = countER - 1
 While countER > 0
    DoCmd.TransferDatabase acImport, "ODBC Database", _
    "ODBC;DSN=DataSource1;UID=" & Decrypt("vtyqx}yx") & ";PWD=" & Decrypt("_|‡pyx€v…A=") & ";LANGUAGE=us_english;" _
    & "DATABASE=pubs", acTable, sqlTBLstr(countER), sqlTBLstr(countER)
    'DoCmd.TransferDatabase acLink, "ODBC Database", "FleetAdvantage_beSQL", acTable, "fa-poseidon.domain.local,1549", sqlTBLstr(countER), True, False
    countER = countER - 1
 Wend
End Function
Public Function WhereAmI(Optional FixConnection As Boolean) As String
    Dim dbs As DataBase
    Dim tdf As TableDef
    Dim Tdfs As TableDefs
    Dim cConnection, cDomain, cTemp As String
    Set dbs = CurrentDb
    Set Tdfs = dbs.TableDefs
    If Environ("USERDOMAIN") = "Domain" Then cDomain = "Office|" Else cDomain = "OutSide|"
    If Environ("ComputerName") = "FA-AZURE-RDS" Then cDomain = "RDS|"

    ' dbo.Units is the canonical system anchor.
    ' If this table is not linked, Atlaas Core is not in a valid state.

    For Each tdf In Tdfs
        If tdf.SourceTableName = "dbo.Units" Then cConnection = tdf.Connect
    Next

    If cConnection Like "*Phoenix*" Then cConnection = "Phoenix"
    If cConnection Like "*10.1.0.7*" Then cConnection = "Poseidon.Local"
    If cConnection Like "*Poseidon.eastus.*" Then cConnection = "Poseidon Web"
    If cConnection Like "*fa-poseidon.domain.local*" Then cConnection = "Poseidon Local"

    If cDomain = "OutSide|" And cConnection = "Phoenix" Then cTemp = "|Mix-OutSide>Phoenix"
    If cDomain = "OutSide|" And cConnection = "Poseidon Local" Then cTemp = "|Mix-OutSide>Local"
    If cDomain = "RDS|" And cConnection = "Poseidon Web" Then cTemp = "|Mix-OutSide>RDS"
    If cDomain = "RDS|" And cConnection = "Phoenix" Then cTemp = "|Mix-Phoenix>RDS"
    If cDomain = "Domain|" And cConnection = "Poseidon Web" Then cTemp = "|Mix-Domain>Web"
    If cDomain = "Office|" And cConnection = "Poseidon Web" Then cTemp = "|Mix-Office>Web"

    WhereAmI = Environ("ComputerName") & "|" & cDomain & cConnection & cTemp

    If FixConnection = True And WhereAmI Like "*Mix*" Then
            'if cTemp =like "Phoenix" then Call ChangeConnectionString("FA-Phoenix", "FleetAdvantage_beSQL", , 0)
             Call ChangeConnectionString
    End If
End Function

Sub ChangeConnectionString(Optional XServerName As String, Optional xDataBase As String, Optional Xtable As String, Optional xUser As Boolean)
  Dim dbs As DataBase
    Dim tdf As TableDef
    Dim Tdfs As TableDefs
    Dim sDriver, UIDType, DataBase As String
    Set dbs = CurrentDb
    Set Tdfs = dbs.TableDefs
    sDriver = "ODBC Driver 17 for SQL Server"
    If Environ("USERDOMAIN") = "DOMAIN" Then UIDType = "Trusted_Connection=yes" Else UIDType = "UID=" & Decrypt("vtyqx}yx") & ";Pwd=" & Decrypt("_|‡pyx€v…A=") & ""

    If xDataBase = "" Then xDataBase = "FleetAdvantage_beSQL"
    If XServerName = "FA-Phoenix" Then xDataBase = "FleetAdvantage_beSQL"
    If XServerName = "" Then
        If Environ("USERDOMAIN") = "DOMAIN" Then XServerName = "10.1.0.7,1549" _
           Else XServerName = "fa-poseidon.eastus.cloudapp.azure.com,1549"
    End If

    If Xtable & "test" = "test" Then
     DoCmd.OpenForm "DataConnection"
     Forms!DataConnection.UserName.caption = "Welcome Back " & DLookup("[Employee Name]", "TBL_Employees", "[User ID] = '" & LCase(Environ("username")) & "'") & vbNewLine & "Just a moment, logging you in."
        For Each tdf In Tdfs    'Loops through list of tables
            'Criteria for selecting which table 'to change the connection string of
            If tdf.SourceTableName Like "*dbo.*" _
               And tdf.Connect <> "" _
               And Not tdf.SourceTableName Like "*TMP*" Then
                If tdf.Connect Like "*Processing*" Then sDataBase = "Processing" Else sDataBase = xDataBase
                tdf.Connect = "ODBC;DRIVER=" & sDriver & ";SERVER=tcp:" & XServerName & ";DATABASE=" & sDataBase & ";" & UIDType & ";Persist Security Info=True;TrustServerCertificate=yes;ConnectPolicy=Redirect;MARS Connection=True;Network Packet Size=16384;"
                DoEvents
                Forms!DataConnection.Message.caption = tdf.SourceTableName
                Forms!DataConnection.Form.Repaint
                ' On Error Resume Next
                If Not tdf.SourceTableName Like "*TMP*" Then tdf.RefreshLink
                On Error GoTo 0
            End If
        Next
    Else
        For Each tdf In Tdfs
            If tdf.SourceTableName = "dbo." & Xtable And tdf.Connect <> "" Then
                tdf.Connect = "ODBC;DRIVER=" & sDriver & ";SERVER=tcp:" & XServerName & ";DATABASE=" & xDataBase & ";" & UIDType & ";Persist Security Info=True;TrustServerCertificate=yes;ConnectPolicy=Redirect;MARS Connection=True;Network Packet Size=16384;"
                'DoCmd.OpenForm "DataConnection"
                'Forms!DataConnection.UserName.caption = "Connecting to Table: " & Xtable & "Just a moment, logging you in."
                tdf.RefreshLink
            End If
        Next
    End If
FixMe:
    'DoCmd.OpenForm "FaMainMenu"
     If IsOpen("FAMainMenu") = True Then Forms!FAMainMenu.DevLabel.caption = DevOrPro
    DoCmd.close acForm, "DataConnection"
End Sub

Function ckenviro()
      Dim i As Integer
        i = 1

        While Environ(i) <> Empty

            Debug.Print Environ(i)

            i = i + 1

        Wend
End Function
