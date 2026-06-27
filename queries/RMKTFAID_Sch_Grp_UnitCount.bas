Operation =1
Option =0
Begin InputTables
    Name ="vw_Sch_Grp_Units_Remarketing"
End
Begin OutputColumns
    Alias ="FAID"
    Expression ="Nz([vw_Sch_Grp_Units_Remarketing].[faid],[ClientShortName] & \"_RMKT\")"
    Expression ="vw_Sch_Grp_Units_Remarketing.Schedule"
    Alias ="Group"
    Expression ="vw_Sch_Grp_Units_Remarketing.UnitGroup"
    Alias ="Units"
    Expression ="Count(vw_Sch_Grp_Units_Remarketing.AssetId)"
    Expression ="vw_Sch_Grp_Units_Remarketing.SGrpID"
    Expression ="vw_Sch_Grp_Units_Remarketing.SchID"
    Expression ="vw_Sch_Grp_Units_Remarketing.ClientShortName"
    Expression ="vw_Sch_Grp_Units_Remarketing.clientID"
End
Begin OrderBy
    Expression ="vw_Sch_Grp_Units_Remarketing.Schedule"
    Flag =0
    Expression ="vw_Sch_Grp_Units_Remarketing.UnitGroup"
    Flag =0
End
Begin Groups
    Expression ="Nz([vw_Sch_Grp_Units_Remarketing].[faid],[ClientShortName] & \"_RMKT\")"
    GroupLevel =0
    Expression ="vw_Sch_Grp_Units_Remarketing.Schedule"
    GroupLevel =0
    Expression ="vw_Sch_Grp_Units_Remarketing.UnitGroup"
    GroupLevel =0
    Expression ="vw_Sch_Grp_Units_Remarketing.SGrpID"
    GroupLevel =0
    Expression ="vw_Sch_Grp_Units_Remarketing.SchID"
    GroupLevel =0
    Expression ="vw_Sch_Grp_Units_Remarketing.ClientShortName"
    GroupLevel =0
    Expression ="vw_Sch_Grp_Units_Remarketing.clientID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FAID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UNITS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.Schedule"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Sch_Grp_Units_Remarketing.clientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =12
    Top =38
    Right =1592
    Bottom =898
    Left =-1
    Top =-1
    Right =1548
    Bottom =441
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =213
        Top =33
        Right =436
        Bottom =422
        Top =0
        Name ="vw_Sch_Grp_Units_Remarketing"
        Name =""
    End
End
