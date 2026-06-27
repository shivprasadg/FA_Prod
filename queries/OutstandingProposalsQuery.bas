Operation =1
Option =0
Where ="(((ClientsView_query.clientID)<>4133) AND ((lookup_transTypes.TransType)=\"Propo"
    "sal\") AND ((Deals.dealLPAFSent) Is Not Null) AND ((Deals.dealoutcomedate) Is Nu"
    "ll) AND ((Groups.AltEquip)=False) AND ((ClientsView_query.ClientGroupID)<> 18))"
Begin InputTables
    Name ="ClientsView_query"
    Name ="Deals"
    Name ="lookup_transTypes"
    Name ="Clients"
    Name ="Groups"
End
Begin OutputColumns
    Expression ="ClientsView_query.clientCompanyName"
    Expression ="ClientsView_query.clientID"
    Expression ="Deals.dealShortDesc"
    Expression ="ClientsView_query.ClientShNm"
    Expression ="Deals.dealID"
    Expression ="ClientsView_query.CLIENTGRPMAX"
    Expression ="ClientsView_query.clientCompanyName"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealUnitQty"
    Expression ="lookup_transTypes.TransType"
    Expression ="Deals.dealGrandTotal"
    Expression ="Groups.groupTotal"
    Expression ="Groups.groupYear"
    Expression ="Groups.groupMake"
    Expression ="Groups.groupModel"
    Expression ="Groups.groupType"
    Expression ="Groups.groupUnits"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealoutcomedate"
    Expression ="Deals.dealoutcome"
    Expression ="Clients.LOinitial"
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
    LeftTable ="Deals"
    RightTable ="Groups"
    Expression ="Deals.dealID = Groups.dealRef"
    Flag =1
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
        dbText "Name" ="Clients_view_query.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3825"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Deals.dealID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Clients_view_query.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFSent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Clients_view_query.CLIENTGRPMAX"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Clients_view_query.ClientShNm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Groups.groupYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Groups.groupMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Groups.groupModel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Groups.groupType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2550"
        dbInteger "ColumnOrder" ="17"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Groups.groupUnits"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Groups.groupTotal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
        dbInteger "ColumnWidth" ="1515"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deals.dealGrandTotal"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deals.faid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1007"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.LOinitial"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealoutcome"
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
    Bottom =216
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =343
        Top =0
        Right =537
        Bottom =265
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =542
        Top =110
        Right =686
        Bottom =254
        Top =0
        Name ="lookup_transTypes"
        Name =""
    End
    Begin
        Left =1
        Top =17
        Right =145
        Bottom =161
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =783
        Top =-14
        Right =1076
        Bottom =257
        Top =0
        Name ="Groups"
        Name =""
    End
End
