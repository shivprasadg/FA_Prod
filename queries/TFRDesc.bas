Operation =1
Option =0
Begin InputTables
    Name ="LookUp_UnitDesc"
End
Begin OutputColumns
    Expression ="LookUp_UnitDesc.GroupID"
    Alias ="ShortDesc"
    Expression ="LookUp_UnitDesc.DescAbbrv"
    Expression ="LookUp_UnitDesc.DescriptionShort"
End
Begin Groups
    Expression ="LookUp_UnitDesc.GroupID"
    GroupLevel =0
    Expression ="LookUp_UnitDesc.DescAbbrv"
    GroupLevel =0
    Expression ="LookUp_UnitDesc.DescriptionShort"
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
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TypeX"
        dbInteger "ColumnWidth" ="5580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ShortDesc"
        dbInteger "ColumnWidth" ="5580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescAbbrv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionShort"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =161
    Top =69
    Right =1315
    Bottom =940
    Left =-1
    Top =-1
    Right =1136
    Bottom =526
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =271
        Top =164
        Right =415
        Bottom =474
        Top =0
        Name ="LookUp_UnitDesc"
        Name =""
    End
End
