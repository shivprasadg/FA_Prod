Operation =1
Option =0
Where ="(((MstrLease.MLNo)>=\"Enter start dt]\" And (MstrLease.MLNo)<=[enter end dt])) O"
    "R (((MstrLease.MLBack)>=\"Enter start dt]\" And (MstrLease.MLBack)<=[enter end d"
    "t]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
End
Begin OutputColumns
    Expression ="MstrLease.MLSent"
    Expression ="Clients.SalespersonREF"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLBack"
    Expression ="Clients.clienttype"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientBecameDate"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLSent"
    Flag =0
    Expression ="Clients.SalespersonREF"
    Flag =0
    Expression ="MstrLease.MLNo"
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
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.SalespersonREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clienttype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="5040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientBecameDate"
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
    Bottom =369
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =443
        Top =22
        Right =733
        Bottom =166
        Top =0
        Name ="MstrLease"
        Name =""
    End
End
