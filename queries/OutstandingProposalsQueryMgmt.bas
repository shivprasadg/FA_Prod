Operation =1
Option =2
Where ="(((Deals.dealoutcomedate) Is Null) AND ((ClientsView_query.clienttype)<>\"client"
    "\") AND ((Clients.ClientShNm) Not Like \"*demo*\"))"
Having ="(((ClientsView_query.clientID)<>4133) AND ((lookup_transTypes.TransType)=\"Propo"
    "sal\") AND ((Deals.dealLPAFSent) Is Not Null))"
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
    Expression ="Deals.dealShortDesc"
    Expression ="ClientsView_query.ClientShNm"
    Expression ="ClientsView_query.CLIENTGRPMAX"
    Expression ="ClientsView_query.clientCompanyName"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealUnitQty"
    Expression ="lookup_transTypes.TransType"
    Expression ="Deals.dealGrandTotal"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealoutcome"
    Expression ="Deals.dealID"
    Expression ="Deals.faid"
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
    Expression ="Deals.dealShortDesc"
    GroupLevel =0
    Expression ="ClientsView_query.ClientShNm"
    GroupLevel =0
    Expression ="ClientsView_query.CLIENTGRPMAX"
    GroupLevel =0
    Expression ="ClientsView_query.clientCompanyName"
    GroupLevel =0
    Expression ="Deals.dealUnitQty"
    GroupLevel =0
    Expression ="lookup_transTypes.TransType"
    GroupLevel =0
    Expression ="Deals.dealGrandTotal"
    GroupLevel =0
    Expression ="Deals.dealLPAFSent"
    GroupLevel =0
    Expression ="Deals.dealoutcome"
    GroupLevel =0
    Expression ="Deals.dealID"
    GroupLevel =0
    Expression ="Deals.faid"
    GroupLevel =0
    Expression ="ClientsView_query.ClientShNm"
    GroupLevel =0
    Expression ="Deals.dealLPAFSent"
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
        dbText "Name" ="Deals.dealUnitQty"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.clientID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFSent"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3300"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Clients_view_query.CLIENTGRPMAX"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.ClientShNm"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealGrandTotal"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.LOinitial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.faid"
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
    Bottom =178
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
