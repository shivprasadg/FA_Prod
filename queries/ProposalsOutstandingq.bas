Operation =1
Option =0
Where ="(((ClientsView_query.clientID)<>4133) And ((Deals.dealLPAFSent) Is Not Null And "
    "(Deals.dealLPAFSent)<=forms!LorPtm!OPAsof) And ((Deals.dealOutcome) Is Null) And"
    " ((lookup_transTypes.TransType)=\"Proposal\") And ((Groups.AltEquip)=No)) Or ((("
    "ClientsView_query.clientID)<>4133) And ((Deals.dealLPAFSent) Is Not Null And (De"
    "als.dealLPAFSent)<=forms!LorPtm!OPAsof) And ((Deals.dealOutcome) Is Not Null) An"
    "d ((Deals.dealOutcomeDate) Is Not Null And (Deals.dealOutcomeDate)>forms!LorPtm!"
    "OPAsof) And ((lookup_transTypes.TransType)<>\"Opportunity\") And ((Groups.AltEqu"
    "ip)=No)) Or (((Deals.dealLPAFSent)>+#1/1/2012# And (Deals.dealLPAFSent)<=#2/29/2"
    "012#))"
Begin InputTables
    Name ="ClientsView_query"
    Name ="Deals"
    Name ="lookup_transTypes"
    Name ="Groups"
    Name ="Salesperson"
End
Begin OutputColumns
    Expression ="ClientsView_query.clientID"
    Expression ="Deals.transTypeRef"
    Expression ="Deals.dealShortDesc"
    Expression ="ClientsView_query.ClientShNm"
    Expression ="Deals.dealID"
    Expression ="ClientsView_query.CLIENTGRPMAX"
    Expression ="ClientsView_query.clientCompanyName"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealUnitQty"
    Expression ="Deals.dealOutcome"
    Expression ="Deals.dealOutcomeDate"
    Expression ="lookup_transTypes.TransType"
    Expression ="Deals.dealGrandTotal"
    Expression ="Groups.groupTotal"
    Expression ="Groups.groupYear"
    Expression ="Groups.groupMake"
    Expression ="Groups.groupModel"
    Expression ="Groups.groupType"
    Expression ="Groups.groupUnits"
    Expression ="Groups.AltEquip"
    Expression ="Salesperson.SPInitials"
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
    LeftTable ="Deals"
    RightTable ="Groups"
    Expression ="Deals.dealID = Groups.dealRef"
    Flag =1
    LeftTable ="ClientsView_query"
    RightTable ="Salesperson"
    Expression ="ClientsView_query.SalespersonREF = Salesperson.SalespersonID"
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
        dbText "Name" ="Deals.transTypeRef"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.clientCompanyName"
        dbInteger "ColumnWidth" ="3825"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealID"
        dbInteger "ColumnOrder" ="5"
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
    End
    Begin
        dbText "Name" ="Clients_view_query.CLIENTGRPMAX"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.ClientShNm"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupYear"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupMake"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupModel"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupType"
        dbInteger "ColumnWidth" ="2550"
        dbInteger "ColumnOrder" ="19"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupUnits"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupTotal"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealGrandTotal"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.AltEquip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SPInitials"
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
    Bottom =365
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
        Left =329
        Top =14
        Right =523
        Bottom =279
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =538
        Top =115
        Right =682
        Bottom =259
        Top =0
        Name ="lookup_transTypes"
        Name =""
    End
    Begin
        Left =756
        Top =12
        Right =1049
        Bottom =283
        Top =0
        Name ="Groups"
        Name =""
    End
    Begin
        Left =48
        Top =240
        Right =322
        Bottom =400
        Top =0
        Name ="Salesperson"
        Name =""
    End
End
