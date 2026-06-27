Attribute VB_Name = "BTF Module"
Option Compare Database
Public Function SendBuyersOrderToBTF(BVDate As Date, BuyersOrderID As Long)
   Dim RST        As DAO.Recordset
   Dim SQLz As String
   Dim AssetIDList As String
   Dim i As Long
    SQLz = "Select UnitID From RMKT_BuyersOrders WHERE BuyersOrderID =" & BuyersOrderID & ";"
    Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
    On Error Resume Next
    MessageUser "Open", "Getting Buyers Order Info", "Please Wait this will take a moment..."
    RST.MoveLast
    RST.MoveFirst
        On Error GoTo 0
   ' On Error GoTo Fixme

    For i = 1 To RST.RecordCount
    AssetIDList = AssetIDList & RST!UnitID & ","
        RST.MoveNext
    Next i
    AssetIDList = Left(AssetIDList, Len(AssetIDList) - 1)
    If AssetIDList <> "" Then Call RunBTF(BVDate, AssetIDList)
FixMe:
    RST.close
    Set RST = Nothing
    On Error GoTo 0
End Function
Public Function RunBTF(BVDate As Date, AssetIDList As String)
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.command
    Dim RST As ADODB.Recordset
    Dim XLS As String
    Dim db As DAO.DataBase
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim i As Long

    '--- Build connection (allows for Sandbox Connections) ---
    '--- cConnection(0) = Server IP;      i.e.: 10.1.0.7,1549;
    '--- cConnection(1) = Connection Type i.e.: Trusted_Connection
    '--- cConnection(2) = DataBase Name   i.e.: FleetAdvantage_beSQL

    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=sqloledb;Server=" & cConnection(0) & _
                            ";Database=" & cConnection(2) & ";" & cConnection(1) & "" & _
                            ";MARS_Connection=yes" & _
                            ";Network Packet Size=16384;"
    conn.Open
    '--- Build command ---
    Set cmd = New ADODB.command
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "dbo.sp_BTF_Base"

        .Parameters.Append .CreateParameter("@BVDate", adDate, adParamInput, , BVDate)
        .Parameters.Append .CreateParameter("@AssetIDList", adVarChar, adParamInput, 6000, AssetIDList)
        .Parameters.Append .CreateParameter("@GroupID", adInteger, adParamInput, , Null)
        .Parameters.Append .CreateParameter("@ClientGroupID", adInteger, adParamInput, , Null)
    End With
    '--- Execute with client-side cursor ---
    Set RST = New ADODB.Recordset
    RST.CursorLocation = adUseClient
    RST.CursorType = adOpenStatic
    RST.LockType = adLockReadOnly
    RST.Open cmd
    If RST.EOF Then
        RST.close
        conn.close
        Exit Function
    End If
    '--- Build DAO temp table ---
    Set db = CurrentDb
    'Delete old table if exists
    On Error Resume Next
    db.TableDefs.Delete "BTFTemp"
    On Error GoTo 0
    'Create new table
    Set tdf = db.CreateTableDef("BTFTemp")
    'Add fields based on ADO metadata
    For i = 0 To RST.fields.Count - 1
        Dim adoFld As ADODB.Field
        Set adoFld = RST.fields(i)

        Select Case adoFld.Type
            Case adInteger, adSmallInt, adTinyInt
                Set fld = tdf.CreateField(adoFld.Name, dbLong)
            Case adBigInt                                           ' SQL Server BIGINT ? DAO Double (safe for all values)
                Set fld = tdf.CreateField(adoFld.Name, dbDouble)
            Case adDouble, adSingle, adFloat
                Set fld = tdf.CreateField(adoFld.Name, dbDouble)
            Case adCurrency
                Set fld = tdf.CreateField(adoFld.Name, dbCurrency)
            Case adNumeric, adDecimal
                Set fld = tdf.CreateField(adoFld.Name, dbDouble)
            Case adBoolean
                Set fld = tdf.CreateField(adoFld.Name, dbBoolean)
            Case adDate, adDBDate, adDBTime, adDBTimeStamp
                Set fld = tdf.CreateField(adoFld.Name, dbDate)
            Case adLongVarChar, adLongVarWChar
                Set fld = tdf.CreateField(adoFld.Name, dbMemo)
            Case adVarChar, adVarWChar, adChar, adWChar            ' Use DefinedSize, cap at 255 for safety
                Dim txtLen As Long
                txtLen = adoFld.DefinedSize
                If txtLen <= 0 Or txtLen > 255 Then
                    Set fld = tdf.CreateField(adoFld.Name, dbMemo)
                Else
                    Set fld = tdf.CreateField(adoFld.Name, dbText, txtLen)
                End If
            Case Else
                Set fld = tdf.CreateField(adoFld.Name, dbText, 255)
        End Select
        tdf.fields.Append fld
    Next i
    db.TableDefs.Append tdf
    '--- Insert rows ---
    Dim rsDAO As DAO.Recordset
    Set rsDAO = db.OpenRecordset("BTFTemp", dbOpenDynaset)

    RST.MoveFirst
    Do Until RST.EOF
        rsDAO.AddNew
        For i = 0 To RST.fields.Count - 1
            If IsNull(RST.fields(i).Value) Then
                rsDAO.fields(RST.fields(i).Name).Value = Null
            Else
                rsDAO.fields(RST.fields(i).Name).Value = RST.fields(i).Value
            End If
        Next i
        rsDAO.Update
        RST.MoveNext
    Loop
    rsDAO.close
    RST.close
    conn.close
    MessageUser "Close"
    MessageUser "Open", "Excel - BTF Automation", "Please Wait this will take a moment..."

          XLS = "|A,mTable:Name=BTF|A,DD3|I,BX|I,cBorder:Med|I,cColor:3:black"
    XLS = XLS & "|C,TC|D,BC|E,TC|E,W10|F,$int|F,W12|H,TC|H,W09|I,TC|J,TC|M,TC|M,BC|M,W07|N,TC|O,TC|O,W06|P,TC|P,W06|Q,TC|R,TC|S,TC|S,W06|T,TC|U,TC|U,W10|V,$E"
    XLS = XLS & "|W,$E|X,$E|Y,$E|Z,$E|Z,W11|AA,$E|AA,W12|AB,$E|AB,W10|AC,$E|AD,%|AD,TC|AD,W06|AE,%|AE,TC|AE,W06|AF,TC|AF,W10|AG,TC|AG,W10|AH,$E|AH,W12"
    XLS = XLS & "|AI,TC|AI,W07|AJ,TC|AJ,W06|AK,TC|AK,P4|AL,%|AL,TC|AL,W07|AM,$E|AM,W12|AN,%|AN,TC|AN,W06|AO,%|AO,TC|AO,W06|AP,$R|AP,TC|AQ,$E|AQ,W12|AR,TC|AR,W10"
    XLS = XLS & "|AS,SintR|AS,W10|AT,SintR|AT,W10|AU,Sint|AU,W10|AV,$E|AV,W10|AW,BC|AW,W06|AX,HH|AY,HH|AZ,HH"
    XLS = XLS & "|V,cBorder:Med:End=AB|AC,cColor:3:green|AC,cBorder:Med"
    XLS = XLS & "|AG,cColor:3:black|AG,cBorder:Med|AH,cColor:3:maroon|AH,cBorder:Med"
    XLS = XLS & "|AI,cColor:3:purple:End=AL|AI,cBorder:Med:End=AL|AI,TC"
    XLS = XLS & "|AM,cColor:3:olive:End=AP|AQ,cColor:3:dark green|AM,cBorder:Med:End=AP|AQ,cBorder:Med"
    XLS = XLS & "|AV,cColor:3:chocolate:End=AV|AV,cBorder:Med|AW,cColor:3:black|"

    SendReport2Excel "SELECT * FROM BTFTemp", "BTF", "BTF For Projected BuyOut Date Of: " & BVDate & "", True, XLS, 0, False
    MessageUser "Close"

End Function
