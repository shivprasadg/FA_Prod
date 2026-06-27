dbMemo "SQL" ="SELECT CustTitleImport.[Record #], vw_SixKeys.AssetID, vw_SixKeys.VIN, CustTitle"
    "Import.[Alternate ID], CustTitleImport.[Loan Amount], CustTitleImport.[Trust Num"
    "ber], vw_SixKeys.CompanyName, vw_SixKeys.FAID, vw_SixKeys.Schedule, vw_SixKeys.U"
    "nitGroup, vw_SixKeys.Status\015\012FROM vw_SixKeys INNER JOIN CustTitleImport ON"
    " vw_SixKeys.VIN = CustTitleImport.VIN\015\012ORDER BY vw_SixKeys.CompanyName, vw"
    "_SixKeys.Schedule, vw_SixKeys.UnitGroup, vw_SixKeys.Status;\015\012"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustTitleImport.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustTitleImport.[Record #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustTitleImport.[Trust Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustTitleImport.[Alternate ID]"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustTitleImport.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustTitleImport.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustTitleImport.[Loan Amount]"
        dbLong "AggregateType" ="-1"
    End
End
