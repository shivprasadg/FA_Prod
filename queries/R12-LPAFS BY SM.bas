Operation =1
Option =0
Where ="(((ClientsView_query.clientID)<>4133) And ((lookup_transTypes.TransType)=\"Propo"
    "sal\") And ((Deals.dealLPAFSent) Is Not Null And (Deals.dealLPAFSent)>=#2/1/2012"
    "# And (Deals.dealLPAFSent)<=#2/29/2012#) And ((ClientsView_query.SPInitials)=\"h"
    "s\" Or (ClientsView_query.SPInitials)=\"sm\") And ((Groups.AltEquip)=No))"
Begin InputTables
    Name ="ClientsView_query"
    Name ="Deals"
    Name ="lookup_transTypes"
    Name ="Groups"
End
Begin OutputColumns
    Expression ="ClientsView_query.clientCompanyName"
    Expression ="ClientsView_query.clientID"
    Alias ="Expr1"
    Expression ="ClientsView_query.SPInitials"
    Expression ="Deals.FAID"
    Expression ="Deals.transTypeRef"
    Expression ="Deals.dealShortDesc"
    Expression ="Deals.dealLPAFSent"
    Expression ="ClientsView_query.ClientShNm"
    Expression ="Deals.dealID"
    Expression ="ClientsView_query.CLIENTGRPMAX"
    Expression ="ClientsView_query.clientCompanyName"
    Alias ="yr"
    Expression ="DatePart(\"yyyy\",[DEALLPAFSENT])"
    Alias ="MO"
    Expression ="DatePart(\"m\",[DEALLPAFSENT])"
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
    Expression ="Deals.FAID"
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
        dbInteger "ColumnOrder" ="12"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealID"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.clientID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFSent"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.CLIENTGRPMAX"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.ClientShNm"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupYear"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupMake"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupModel"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupType"
        dbInteger "ColumnWidth" ="2550"
        dbInteger "ColumnOrder" ="21"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupUnits"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupTotal"
        dbInteger "ColumnWidth" ="1515"
        dbInteger "ColumnOrder" ="16"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealGrandTotal"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="yr"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MO"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
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
    Bottom =-1
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
        Left =609
        Top =94
        Right =753
        Bottom =238
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
End
