Operation =1
Option =0
Where ="(((Units.UnitUnitNum) In (SELECT [UnitUnitNum] FROM [Units] As Tmp GROUP BY [Uni"
    "tUnitNum],[ClientGroupId] HAVING Count(*)>1  And [ClientGroupId] = [Units].[Clie"
    "ntGroupId])) AND ((Units.ClientGroupId)<>91 And (Units.ClientGroupId)<>53 And (U"
    "nits.ClientGroupId)<>50))"
Begin InputTables
    Name ="Units"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="Units.UnitUnitNum"
    Expression ="Units.ClientGroupId"
    Expression ="Units.UnitID"
    Expression ="Units.GroupREF"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitVIN"
    Expression ="Units.clientID"
    Expression ="Units.UnitSubType"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitId"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitUnitNum"
    Flag =0
    Expression ="Units.ClientGroupId"
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
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitStatus]"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitUnitNum]"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[GroupREF]"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[ClientGroupId]"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.clientID"
        dbInteger "ColumnWidth" ="1140"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-404
    Top =28
    Right =1114
    Bottom =1009
    Left =-1
    Top =-1
    Right =1486
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
        Left =244
        Top =65
        Right =388
        Bottom =268
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
End
