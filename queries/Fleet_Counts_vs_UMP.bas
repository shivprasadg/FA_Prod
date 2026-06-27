dbMemo "SQL" ="SELECT Min(vw_SixKeys.ClientGroupId) AS ClientGroupId, vw_SixKeys.clientID, vw_S"
    "ixKeys.CompanyName, vw_SixKeys.Status, Count(vw_SixKeys.Unitid) AS Units, Count("
    "UMP_DataCount.UnitId) AS HasOBCData, [Units]-[HasOBCData] AS NoOBCData, vw_SixKe"
    "ys.CompanyType\015\012FROM vw_SixKeys LEFT JOIN UMP_DataCount ON vw_SixKeys.Unit"
    "ID = UMP_DataCount.UnitId\015\012GROUP BY vw_SixKeys.clientID, vw_SixKeys.Compan"
    "yName, vw_SixKeys.Status, vw_SixKeys.CompanyType;\015\012"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoOBCData"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HasOBCData"
        dbLong "AggregateType" ="-1"
    End
End
