Operation =1
Option =0
Begin InputTables
    Name ="LookUp_UnitDesc"
End
Begin OutputColumns
    Expression ="LookUp_UnitDesc.GroupID"
    Alias ="Units"
    Expression ="Count(LookUp_UnitDesc.AssetID)"
    Alias ="UnitDesc"
    Expression ="\"(\" & Count([assetid]) & \") \" & IIf(Len([LegalDescription])<3,[DescriptionLo"
        "ng],[LegalDescription])"
    Expression ="LookUp_UnitDesc.DescriptionLong"
    Expression ="LookUp_UnitDesc.LegalDescription"
End
Begin Groups
    Expression ="LookUp_UnitDesc.GroupID"
    GroupLevel =0
    Expression ="LookUp_UnitDesc.DescriptionLong"
    GroupLevel =0
    Expression ="LookUp_UnitDesc.LegalDescription"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionLong"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.LegalDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbInteger "ColumnWidth" ="5355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_UnitDesc.DescriptionShort"
        dbInteger "ColumnWidth" ="4260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =133
    Top =110
    Right =1304
    Bottom =1087
    Left =-1
    Top =-1
    Right =1153
    Bottom =649
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =401
        Top =142
        Right =545
        Bottom =526
        Top =0
        Name ="LookUp_UnitDesc"
        Name =""
    End
End
