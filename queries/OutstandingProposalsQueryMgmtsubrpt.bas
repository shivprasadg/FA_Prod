Operation =1
Option =2
Where ="(((Deals.dealLPAFSent) Is Not Null) AND ((lookup_transTypes.TransType)=\"Proposa"
    "l\") AND ((Deals.dealoutcomedate) Is Null) AND ((ClientsView_query.clienttype)<>"
    "\"client\") AND ((Clients.ClientShNm) Not Like \"*demo*\"))"
Having ="(((ClientsView_query.clientID)<>4133))"
Begin InputTables
    Name ="ClientsView_query"
    Name ="Deals"
    Name ="lookup_transTypes"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="ClientsView_query.ClientShNm"
    Expression ="ClientsView_query.clientID"
    Expression ="Clients.LOinitial"
    Expression ="Deals.transTypeRef"
    Expression ="ClientsView_query.ClientShNm"
    Expression ="ClientsView_query.CLIENTGRPMAX"
    Expression ="ClientsView_query.clientCompanyName"
    Alias ="MinOfdealLPAFSent"
    Expression ="Min(Deals.dealLPAFSent)"
End
Begin Joins
    LeftTable ="ClientsView_query"
    RightTable ="Deals"
    Expression ="ClientsView_query.clientID = Deals.clientRef"
    Flag =1
    LeftTable ="Deals"
    RightTable ="lookup_transTypes"
    Expression ="Deals.transTypeRef = lookup_transTypes.TransTypeID"
    Flag =1
    LeftTable ="ClientsView_query"
    RightTable ="Clients"
    Expression ="ClientsView_query.clientID = Clients.clientID"
    Flag =1
End
Begin Groups
    Expression ="ClientsView_query.clientID"
    GroupLevel =0
    Expression ="Clients.LOinitial"
    GroupLevel =0
    Expression ="Deals.transTypeRef"
    GroupLevel =0
    Expression ="ClientsView_query.ClientShNm"
    GroupLevel =0
    Expression ="ClientsView_query.CLIENTGRPMAX"
    GroupLevel =0
    Expression ="ClientsView_query.clientCompanyName"
    GroupLevel =0
    Expression ="ClientsView_query.ClientShNm"
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
        dbText "Name" ="Deals.transTypeRef"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.clientCompanyName"
        dbInteger "ColumnWidth" ="3825"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.clientID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.CLIENTGRPMAX"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.ClientShNm"
        dbInteger "ColumnWidth" ="1830"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.LOinitial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfdealLPAFSent"
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
    Bottom =193
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ClientsView_query"
        Name =""
    End
    Begin
        Left =446
        Top =-3
        Right =640
        Bottom =262
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =700
        Top =94
        Right =844
        Bottom =238
        Top =0
        Name ="lookup_transTypes"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
End
