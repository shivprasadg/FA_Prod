Operation =1
Option =0
Where ="(((vw_sixkeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_sixkeys"
End
Begin OutputColumns
    Expression ="vw_sixkeys.ClientShortName"
    Alias ="AYear"
    Expression ="Min(DatePart(\"yyyy\",[AcceptedDate]))"
    Alias ="UnitAccepted"
    Expression ="Count(vw_sixkeys.AssetID)"
    Expression ="vw_sixkeys.clientID"
    Expression ="vw_sixkeys.ClientGroupId"
    Alias ="FristAccepted"
    Expression ="Min(vw_sixkeys.AcceptedDate)"
    Alias ="LastAccepted"
    Expression ="Max(vw_sixkeys.AcceptedDate)"
End
Begin OrderBy
    Expression ="Min(DatePart(\"yyyy\",[AcceptedDate]))"
    Flag =1
End
Begin Groups
    Expression ="vw_sixkeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_sixkeys.clientID"
    GroupLevel =0
    Expression ="vw_sixkeys.ClientGroupId"
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
        dbText "Name" ="UnitAccepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastAccepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FristAccepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_sixkeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_sixkeys.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_sixkeys.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AYear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =165
    Top =27
    Right =1592
    Bottom =1025
    Left =-1
    Top =-1
    Right =1409
    Bottom =612
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_sixkeys"
        Name =""
    End
End
