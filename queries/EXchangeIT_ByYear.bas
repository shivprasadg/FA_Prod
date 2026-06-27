Operation =1
Option =0
Having ="(((UnitExitOption.OptionTypeID)=1))"
Begin InputTables
    Name ="UnitExitOption"
End
Begin OutputColumns
    Expression ="UnitExitOption.ReplacedByGroupID"
    Expression ="UnitExitOption.OptionTypeID"
    Expression ="UnitExitOption.OptionTypeDesc"
    Alias ="Units"
    Expression ="Count(UnitExitOption.UnitID)"
End
Begin OrderBy
    Expression ="UnitExitOption.OptionTypeID"
    Flag =0
End
Begin Groups
    Expression ="UnitExitOption.ReplacedByGroupID"
    GroupLevel =0
    Expression ="UnitExitOption.OptionTypeID"
    GroupLevel =0
    Expression ="UnitExitOption.OptionTypeDesc"
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
        dbText "Name" ="UnitExitOption.OptionTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="0"
    End
End
Begin
    State =2
    Left =-8
    Top =-31
    Right =1332
    Bottom =994
    Left =-1
    Top =-1
    Right =1308
    Bottom =669
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =275
        Top =118
        Right =510
        Bottom =514
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
End
