Operation =1
Option =2
Begin InputTables
    Name ="vw_ATF_Avail_Signed_FAID"
    Name ="Clients"
    Name ="Salesperson"
End
Begin OutputColumns
    Expression ="vw_ATF_Avail_Signed_FAID.ClientGroupId"
    Expression ="vw_ATF_Avail_Signed_FAID.clientID"
    Alias ="BDEX"
    Expression ="Salesperson.SalespersonID"
    Alias ="LesseeID"
    Expression ="vw_ATF_Avail_Signed_FAID.clientID"
    Alias ="Lessee"
    Expression ="Clients.clientCompanyName"
    Alias ="TaxID"
    Expression ="Clients.clientTaxID"
    Alias ="LOExec"
    Expression ="vw_ATF_Avail_Signed_FAID.BDEX"
    Expression ="vw_ATF_Avail_Signed_FAID.FAID"
    Expression ="vw_ATF_Avail_Signed_FAID.SchName"
    Alias ="SchShortDesc"
    Expression ="vw_ATF_Avail_Signed_FAID.Schedule_SchShortDesc"
    Alias ="DealUnitQty"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealUnitQty"
    Expression ="vw_ATF_Avail_Signed_FAID.ShortGUID"
End
Begin Joins
    LeftTable ="vw_ATF_Avail_Signed_FAID"
    RightTable ="Clients"
    Expression ="vw_ATF_Avail_Signed_FAID.clientID = Clients.clientID"
    Flag =2
    LeftTable ="vw_ATF_Avail_Signed_FAID"
    RightTable ="Salesperson"
    Expression ="vw_ATF_Avail_Signed_FAID.BDEX = Salesperson.SPInitials"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.clientCompanyName"
    Flag =0
    Expression ="vw_ATF_Avail_Signed_FAID.FAID"
    Flag =1
End
Begin Groups
    Expression ="vw_ATF_Avail_Signed_FAID.ClientGroupId"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.clientID"
    GroupLevel =0
    Expression ="Salesperson.SalespersonID"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.clientID"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Clients.clientTaxID"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.BDEX"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.FAID"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.SchName"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Schedule_SchShortDesc"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealUnitQty"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.ShortGUID"
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
        dbText "Name" ="SchShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealUnitQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BDEX"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LOExec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LesseeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.SchName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.ShortGUID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3615"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =-310
    Top =138
    Right =1289
    Bottom =616
    Left =-1
    Top =-1
    Right =1581
    Bottom =113
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_ATF_Avail_Signed_FAID"
        Name =""
    End
    Begin
        Left =237
        Top =69
        Right =381
        Bottom =213
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =447
        Top =18
        Right =641
        Bottom =204
        Top =0
        Name ="Salesperson"
        Name =""
    End
End
