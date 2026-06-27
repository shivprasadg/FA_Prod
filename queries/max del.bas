Operation =1
Option =0
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.GroupREF"
    Alias ="Maxin-serv"
    Expression ="Max(IIf(IsNull([unitinserv]),IIf(IsNull([unitestdeldate]),[unitbuilddate]),[unit"
        "inserv]))"
    Alias ="ProjSyn"
    Expression ="Max(DateAdd(\"m\",1,IIf(IsNull([unitinserv]),IIf(IsNull([unitestdeldate]),[unitb"
        "uilddate]),[unitinserv]))-1)"
    Alias ="Deldt"
    Expression ="Max(IIf(IsNull([unitactualdeldate]),[unitestdeldate],[unitactualdeldate]))"
    Alias ="MaxAccpt"
    Expression ="Max(Units.UnitAcceptDate)"
End
Begin Groups
    Expression ="Units.GroupREF"
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
        dbText "Name" ="ProjSyn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Maxin-serv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deldt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxAccpt"
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
    Bottom =106
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =287
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
End
