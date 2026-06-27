dbMemo "SQL" ="SELECT S.BDEX, S.ClientShortName AS Account, PD.AssetID, PD.BLCD, PD.PerdiemStar"
    "tDate AS PerDiemStart, PD.PerDiemDays, PD.RentPD AS PerDiemRent, PD.PerDiemAmt, "
    "PD.PerDiemProjected AS PerDiemProj\015\012FROM vw_SixKeys AS S INNER JOIN PerDie"
    "m_Days_Rate_Unit AS PD ON S.UnitID = PD.AssetID\015\012ORDER BY S.BDEX, S.Client"
    "ShortName;\015\012"
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
dbMemo "Filter" ="([PerDiem_Days_Rate_BDEX].[PerDiemStart] Is Null)"
Begin
    Begin
        dbText "Name" ="PerDiemProj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiem_Days_Rate_Unit.PerDiemDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiem_Days_Rate_Unit.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiem_Days_Rate_Unit.PerDiemAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiemRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Account"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiem_Days_Rate_Unit.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiemStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PD.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PD.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PD.PerDiemDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PD.PerDiemAmt"
        dbLong "AggregateType" ="-1"
    End
End
