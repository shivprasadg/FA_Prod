Operation =1
Option =0
Where ="(((Units.UnitStatus)=\"A\" Or (Units.UnitStatus)=\"B\" Or (Units.UnitStatus)=\"C"
    "\" Or (Units.UnitStatus)=\"O\" Or (Units.UnitStatus)=\"S\"))"
Having ="(((vw_SixKeys.MLOrig)=\"fa\"))"
Begin InputTables
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="Units.GroupREF"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & IIf(Nz([UnitSubType],\"\")"
        "=\"\",\"\",\" \" & [UnitSubType]) & \" \" & [UnitType]"
    Alias ="Units"
    Expression ="Count(Units.UnitID)"
    Expression ="vw_SixKeys.MLOrig"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
    Flag =1
End
Begin Groups
    Expression ="Units.GroupREF"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & IIf(Nz([UnitSubType],\"\")"
        "=\"\",\"\",\" \" & [UnitSubType]) & \" \" & [UnitType]"
    GroupLevel =0
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
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="4905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =253
    Top =66
    Right =1680
    Bottom =899
    Left =-1
    Top =-1
    Right =1395
    Bottom =401
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =118
        Top =89
        Right =333
        Bottom =442
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =404
        Top =127
        Right =548
        Bottom =406
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
