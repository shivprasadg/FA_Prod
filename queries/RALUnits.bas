Operation =1
Option =0
Begin InputTables
    Name ="RALTemp1"
    Name ="vw_RALAttachment"
End
Begin OutputColumns
    Expression ="RALTemp1.Letterdate"
    Expression ="vw_RALAttachment.*"
    Alias ="GroupsIn"
    Expression ="GroupsNoinSchedule([vw_RALAttachment].[SchID])"
End
Begin Joins
    LeftTable ="RALTemp1"
    RightTable ="vw_RALAttachment"
    Expression ="RALTemp1.SchID = vw_RALAttachment.SchID"
    Flag =1
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
dbInteger "RowHeight" ="660"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="GroupsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbInteger "ColumnWidth" ="4995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALTemp1.Letterdate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LessorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.UnitDescDocVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.LAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.GroupNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.LessorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.UnitRentSynd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.DelvDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.FundedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_RALAttachment.SGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =33
    Top =304
    Right =1477
    Bottom =1168
    Left =-1
    Top =-1
    Right =1420
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-5
        Top =11
        Right =139
        Bottom =242
        Top =0
        Name ="RALTemp1"
        Name =""
    End
    Begin
        Left =353
        Top =36
        Right =497
        Bottom =180
        Top =0
        Name ="vw_RALAttachment"
        Name =""
    End
End
