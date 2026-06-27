Operation =1
Option =0
Begin InputTables
    Name ="RALTemp1"
    Name ="vw_SixKeys"
    Name ="UnitTotalOEC"
End
Begin OutputColumns
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.UnitID)"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="Lessorcost"
    Expression ="Sum(UnitTotalOEC.UnitTotalCost)"
    Alias ="Desc"
    Expression ="UnitDescinGroup([vw_sixkeys].[sGrpid],True)"
    Alias ="SyndRent"
    Expression ="UnitRentsinGroup(vw_SixKeys.[SGrpID],True)"
    Alias ="BLCD"
    Expression ="Format([vw_sixkeys].[BLCD],\"mmmm dd\"\", \"\"yyyy\")"
    Expression ="vw_SixKeys.LeaseTermPrimary"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="UnitTotalOEC"
    Expression ="vw_SixKeys.AssetID = UnitTotalOEC.AssetID"
    Flag =1
    LeftTable ="RALTemp1"
    RightTable ="vw_SixKeys"
    Expression ="RALTemp1.SchID = vw_SixKeys.SchID"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.SchID"
    GroupLevel =0
    Expression ="vw_SixKeys.SGrpID"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitGroup"
    GroupLevel =0
    Expression ="Format([vw_sixkeys].[BLCD],\"mmmm dd\"\", \"\"yyyy\")"
    GroupLevel =0
    Expression ="vw_SixKeys.LeaseTermPrimary"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbInteger "RowHeight" ="1935"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndRent"
        dbInteger "ColumnWidth" ="3135"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessorcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =246
    Top =76
    Right =1869
    Bottom =1070
    Left =-1
    Top =-1
    Right =1599
    Bottom =632
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =27
        Top =71
        Right =171
        Bottom =458
        Top =0
        Name ="RALTemp1"
        Name =""
    End
    Begin
        Left =238
        Top =40
        Right =499
        Bottom =699
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =706
        Top =131
        Right =850
        Bottom =450
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
End
