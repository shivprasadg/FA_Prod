Operation =1
Option =0
Begin InputTables
    Name ="UnitExitOption"
End
Begin OutputColumns
    Expression ="UnitExitOption.UnitID"
    Alias ="Exch"
    Expression ="IIf([OptionTypeID]=1,\"Exch\",Null)"
    Alias ="Term"
    Expression ="IIf([OptionTypeID]=2,\"Term\",Null)"
    Alias ="Extend"
    Expression ="IIf([OptionTypeID]>2,\"Extd\",Null)"
End
Begin Groups
    Expression ="UnitExitOption.UnitID"
    GroupLevel =0
    Expression ="IIf([OptionTypeID]=1,\"Exch\",Null)"
    GroupLevel =0
    Expression ="IIf([OptionTypeID]=2,\"Term\",Null)"
    GroupLevel =0
    Expression ="IIf([OptionTypeID]>2,\"Extd\",Null)"
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
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Term"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Extend"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =232
    Top =161
    Right =1560
    Bottom =1042
    Left =-1
    Top =-1
    Right =1296
    Bottom =546
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =73
        Top =80
        Right =388
        Bottom =330
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
End
