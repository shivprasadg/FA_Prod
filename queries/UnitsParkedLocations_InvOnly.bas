Operation =1
Option =0
Where ="(((IIf([status]=\"A\" Or [status]=\"B\" Or [status]=\"R\" Or [status]=\"T\" Or ["
    "status]=\"TR\",1,0))=1))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Alias ="GarageLocationID"
    Expression ="vw_SixKeys.LocationId"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitNum"
    Alias ="Desc"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Expression ="vw_SixKeys.MYear"
    Expression ="vw_SixKeys.Make"
    Expression ="vw_SixKeys.Model"
    Expression ="vw_SixKeys.Type"
    Expression ="vw_SixKeys.SubType"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.ClientID"
    Expression ="vw_SixKeys.MLOrig"
    Alias ="InFleet"
    Expression ="IIf([status]=\"A\" Or [status]=\"B\" Or [status]=\"R\" Or [status]=\"T\" Or [sta"
        "tus]=\"TR\",1,0)"
End
Begin OrderBy
    Expression ="vw_SixKeys.MLOrig"
    Flag =0
    Expression ="vw_SixKeys.UnitID"
    Flag =0
End
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
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GarageLocationID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="4530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InFleet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescShortVer"
        dbInteger "ColumnWidth" ="4530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitNum"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =86
    Top =43
    Right =1348
    Bottom =977
    Left =-1
    Top =-1
    Right =1244
    Bottom =344
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =437
        Top =12
        Right =581
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
