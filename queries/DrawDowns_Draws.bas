dbMemo "SQL" ="SELECT DrawDowns.DrawDownId, Banks.BankShortName, DrawDowns.DrawRequestNumber, D"
    "rawDowns.DrawNumber, DrawDowns.DrawName, Count(vw_SixKeys.AssetID) AS Units, Cou"
    "nt(ChildParts.ChildPartID) AS Parts, Sum(vw_SixKeys.UnitCost) AS UnitAmt, Sum(Ch"
    "ildParts.childFinalCost) AS PartAmt\015\012FROM ((DrawDowns LEFT JOIN vw_SixKeys"
    " ON DrawDowns.DrawDownId = vw_SixKeys.DrawDownId) LEFT JOIN ChildParts ON DrawDo"
    "wns.DrawDownId = ChildParts.DrawDownId) INNER JOIN Banks ON DrawDowns.DrawBankId"
    " = Banks.BankId\015\012WHERE (((DrawDowns.MLID)<>0))\015\012GROUP BY DrawDowns.D"
    "rawDownId, Banks.BankShortName, DrawDowns.DrawRequestNumber, DrawDowns.DrawNumbe"
    "r, DrawDowns.DrawName\015\012ORDER BY DrawDowns.DrawRequestNumber, DrawDowns.Dra"
    "wNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawRequestNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartAmt"
        dbLong "AggregateType" ="-1"
    End
End
