Operation =1
Option =0
Begin InputTables
    Name ="LeaseInOutExtendedBase"
End
Begin OutputColumns
    Alias ="UnitsOnLease2015"
    Expression ="Sum(LeaseInOutExtendedBase.OnLease2015)"
    Alias ="UnitsAdded2015"
    Expression ="Sum(LeaseInOutExtendedBase.Added2015)"
    Alias ="UnitsTerminated2015"
    Expression ="Sum(LeaseInOutExtendedBase.Terminated2015)"
    Alias ="UnitsExtended2015"
    Expression ="Sum(LeaseInOutExtendedBase.Extended2015)"
    Alias ="OnLease2016"
    Expression ="([UnitsOnLease2015]+[UnitsAdded2015])-[UnitsTerminated2015]"
    Alias ="Added2016"
    Expression ="Sum(LeaseInOutExtendedBase.Added2016)"
    Alias ="Terminated2016"
    Expression ="Sum(LeaseInOutExtendedBase.Terminated2016)"
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
        dbText "Name" ="Terminated2016"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Added2016"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="UnitsOnLease2015"
        dbInteger "ColumnWidth" ="2115"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsAdded2015"
        dbInteger "ColumnWidth" ="1935"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnLease2016"
        dbInteger "ColumnWidth" ="1635"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsExtended2015"
        dbInteger "ColumnWidth" ="1725"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsTerminated2015"
        dbInteger "ColumnWidth" ="2400"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-169
    Top =121
    Right =1514
    Bottom =997
    Left =-1
    Top =-1
    Right =1665
    Bottom =582
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =136
        Top =35
        Right =280
        Bottom =333
        Top =0
        Name ="LeaseInOutExtendedBase"
        Name =""
    End
End
