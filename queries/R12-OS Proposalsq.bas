Operation =1
Option =0
Where ="(((Deals.dealoutcomedate) Is Null) And ((lookup_transTypes.TransType)=\"Proposal"
    "\") And ((Deals.dealLPAFSent) Is Not Null And (Deals.dealLPAFSent)<=#3/31/2012#)"
    " And ((ClientsView_query.SPInitials)=\"sm\" Or (ClientsView_query.SPInitials)=\""
    "hs\") And ((ClientsView_query.clientID)<>4133) And ((Groups.AltEquip)=No)) Or (("
    "(Deals.dealoutcomedate) Is Not Null And (Deals.dealoutcomedate)>#2/29/2012#) And"
    " ((lookup_transTypes.TransType)=\"Proposal\") And ((Deals.dealLPAFSent) Is Not N"
    "ull And (Deals.dealLPAFSent)<=#3/31/2012#) And ((ClientsView_query.SPInitials)=\""
    "sm\" Or (ClientsView_query.SPInitials)=\"hs\") And ((ClientsView_query.clientID)"
    "<>4133) And ((Groups.AltEquip)=No))"
Begin InputTables
    Name ="ClientsView_query"
    Name ="Deals"
    Name ="lookup_transTypes"
    Name ="Groups"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="ClientsView_query.SPInitials"
    Expression ="ClientsView_query.clientCompanyName"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealoutcomedate"
    Expression ="Deals.transTypeRef"
    Expression ="Deals.dealShortDesc"
    Expression ="ClientsView_query.ClientShNm"
    Expression ="Deals.dealID"
    Expression ="ClientsView_query.CLIENTGRPMAX"
    Expression ="ClientsView_query.clientCompanyName"
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
    Expression ="Deals.dealoutcome"
    Expression ="ClientsView_query.clientshnm"
End
Begin Joins
    LeftTable ="ClientsView_query"
    RightTable ="Deals"
    Expression ="ClientsView_query.clientID = Deals.clientRef"
    Flag =1
    LeftTable ="Deals"
    RightTable ="lookup_transTypes"
    Expression ="Deals.transTypeRef = lookup_transTypes.TransTypeID"
    Flag =2
    LeftTable ="Deals"
    RightTable ="Groups"
    Expression ="Deals.dealID = Groups.dealRef"
    Flag =1
End
Begin OrderBy
    Expression ="ClientsView_query.SPInitials"
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
        dbText "Name" ="Deals.transTypeRef"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.clientCompanyName"
        dbInteger "ColumnWidth" ="3825"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealID"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFSent"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.CLIENTGRPMAX"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.ClientShNm"
        dbInteger "ColumnWidth" ="2220"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupYear"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupMake"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupModel"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupType"
        dbInteger "ColumnWidth" ="2550"
        dbInteger "ColumnOrder" ="20"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupUnits"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupTotal"
        dbInteger "ColumnWidth" ="1515"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealGrandTotal"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Expr1006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1247
    Bottom =679
    Left =-1
    Top =-1
    Right =1215
    Bottom =339
    Left =768
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
        Left =-440
        Top =-6
        Right =-246
        Bottom =259
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =36
        Top =-13
        Right =180
        Bottom =131
        Top =0
        Name ="lookup_transTypes"
        Name =""
    End
    Begin
        Left =-108
        Top =117
        Right =185
        Bottom =388
        Top =0
        Name ="Groups"
        Name =""
    End
End
