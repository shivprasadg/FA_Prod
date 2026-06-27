Operation =1
Option =0
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.Make"
    Expression ="vw_SixKeys.Model"
    Alias ="Leased"
    Expression ="IIf([vw_SixKeys].[MLOrig]=\"FA\",Count([UnitID]),0)"
    Alias ="Tracked"
    Expression ="IIf([vw_SixKeys].[MLOrig]=\"FA\",0,Count([UnitID]))"
    Expression ="vw_SixKeys.MLOrig"
End
Begin OrderBy
    Expression ="vw_SixKeys.Make"
    Flag =0
    Expression ="vw_SixKeys.Model"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.Make"
    GroupLevel =0
    Expression ="vw_SixKeys.Model"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
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
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tracked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Model"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Leased"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1430
    Bottom =991
    Left =-1
    Top =-1
    Right =1354
    Bottom =597
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =352
        Top =12
        Right =642
        Bottom =366
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
