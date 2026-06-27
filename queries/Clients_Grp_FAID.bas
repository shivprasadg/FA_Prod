Operation =1
Option =0
Begin InputTables
    Name ="Clients_FAID"
    Name ="Clients_FAID"
    Alias ="Clients_FAID_1"
End
Begin OutputColumns
    Expression ="Clients_FAID.FAID"
    Alias ="ClientShNm"
    Expression ="First(Clients_FAID_1.ClientShNm)"
    Alias ="clientID"
    Expression ="First(Clients_FAID_1.clientID)"
End
Begin Joins
    LeftTable ="Clients_FAID"
    RightTable ="Clients_FAID_1"
    Expression ="Clients_FAID.FAID = Clients_FAID_1.FAID"
    Flag =1
End
Begin OrderBy
    Expression ="Clients_FAID.FAID"
    Flag =0
    Expression ="First(Clients_FAID_1.ClientShNm)"
    Flag =0
End
Begin Groups
    Expression ="Clients_FAID.FAID"
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
        dbText "Name" ="Clients_FAID.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientShNm"
        dbInteger "ColumnWidth" ="3765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1358
    Bottom =855
    Left =-1
    Top =-1
    Right =1326
    Bottom =481
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =175
        Top =104
        Right =319
        Bottom =248
        Top =0
        Name ="Clients_FAID"
        Name =""
    End
    Begin
        Left =404
        Top =159
        Right =548
        Bottom =303
        Top =0
        Name ="Clients_FAID_1"
        Name =""
    End
End
