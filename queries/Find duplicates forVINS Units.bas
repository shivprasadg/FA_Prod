Operation =1
Option =0
Where ="(((Units.UnitVIN) In (SELECT [UnitVIN] FROM [Units] As Tmp GROUP BY [UnitVIN] HA"
    "VING Count(*)>1 )))"
Begin InputTables
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.SGrpID"
    Expression ="vw_SixKeys.clientID"
    Expression ="vw_SixKeys.ClientGroupId"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitID"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitSubType"
    Expression ="Units.ClientGroupId"
    Expression ="Units.ClientID"
    Expression ="Units.UnitStatus"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =1
    Expression ="Units.UnitVIN"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([Find duplicates forVINS Units].[UnitStatus]=\"T\")"
dbMemo "OrderBy" ="[Find duplicates forVINS Units].[UnitVIN], [vw_SixKeys].[ClientGroupId] DESC"
Begin
    Begin
        dbText "Name" ="Units.[UnitModel]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitVIN]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitSubType]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[ClientGroupId]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.[UnitUnitNum]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[ClientID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitYr]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitMake]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitStatus]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =67
    Top =127
    Right =1762
    Bottom =1025
    Left =-1
    Top =-1
    Right =1558
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
