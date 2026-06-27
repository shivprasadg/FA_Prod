Operation =1
Option =0
Begin InputTables
    Name ="HrsOnlyq"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="HrsOnlyq.Unitref"
    Alias ="MaxOfReefer Hours"
    Expression ="Max(HrsOnlyq.[Reefer Hours])"
    Alias ="MaxOfClose DAte"
    Expression ="Max(HrsOnlyq.[Close DAte])"
End
Begin Groups
    Expression ="HrsOnlyq.Unitref"
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
        dbText "Name" ="MaxOfReefer Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfClose DAte"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HrsOnlyq.Unitref"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =306
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="HrsOnlyq"
        Name =""
    End
End
