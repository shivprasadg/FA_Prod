Operation =3
Name ="FAScRptT1"
Option =2
Where ="(((Clients.ClientShNm)<>\"test\") AND ((Deals.dealOutcome) Is Null And (Deals.de"
    "alOutcome) Is Null) AND ((Deals.transTypeRef)=2)) OR (((Clients.ClientShNm)<>\"t"
    "est\") AND ((Deals.dealOutcome) Is Null And (Deals.dealOutcome) Is Null) AND ((D"
    "eals.transTypeRef)=3))"
Begin InputTables
    Name ="Clients"
    Name ="Salesperson"
    Name ="Deals"
    Name ="Groups"
End
Begin OutputColumns
    Name ="ClientShNm"
    Expression ="Clients.ClientShNm"
    Name ="clientCompanyName"
    Expression ="Clients.clientCompanyName"
    Name ="creditFAProfile"
    Expression ="Clients.creditFAProfile"
    Name ="ClientGroupID"
    Expression ="Clients.ClientGroupID"
    Name ="creditFAScore"
    Expression ="Clients.creditFAScore"
    Name ="SPInitials"
    Expression ="Salesperson.SPInitials"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Salesperson"
    Expression ="Clients.SalespersonREF = Salesperson.SalespersonID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Clients_view_query.creditFAProfile"
        dbInteger "ColumnWidth" ="1965"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.clientCompanyName"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4680"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Clients_view_query.JFApproval"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.creditFAScore"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.ClientShNm"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Pipeline"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.[client Group]"
        dbInteger "ColumnWidth" ="2865"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_view_query.prospecttopten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAProfile"
        dbInteger "ColumnWidth" ="1965"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="3495"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.prospecttopten"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[client group]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.transTypeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAScore"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Top45"
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
    Bottom =378
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =-176
        Top =4
        Right =54
        Bottom =344
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =204
        Top =216
        Right =444
        Bottom =360
        Top =0
        Name ="Salesperson"
        Name =""
    End
    Begin
        Left =205
        Top =1
        Right =445
        Bottom =195
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =515
        Top =2
        Right =659
        Bottom =180
        Top =0
        Name ="Groups"
        Name =""
    End
End
