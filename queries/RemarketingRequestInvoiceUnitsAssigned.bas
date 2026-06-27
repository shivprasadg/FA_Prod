Operation =1
Option =0
Where ="(((Rmkt.InvoiceRequestNum) Is Not Null))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Rmkt"
    Name ="Clients"
End
Begin OutputColumns
    Alias ="UnitID"
    Expression ="Rmkt.UnitRef"
    Alias ="Num"
    Expression ="Rmkt.InvoiceRequestNum"
    Alias ="Requested"
    Expression ="Format([InvoiceRequestDate],\"mm/dd/yy\")"
    Alias ="Rep"
    Expression ="Rmkt.RmkSm"
    Expression ="vw_SixKeys.AssetID"
    Alias ="Unit#"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.VIN"
    Alias ="Description"
    Expression ="[MYear] & \" \" & [Make] & \" \" & [Model] & IIf(Nz([SubType],\"\")=\"\",\"\",\""
        " \" & [SubType]) & IIf([Type]=\"Tractor\",\"\",\" \" & [Type])"
    Expression ="vw_SixKeys.Status"
    Alias ="Sold Price"
    Expression ="Rmkt.SoldPrice"
    Expression ="Rmkt.InvoiceSentDate"
    Alias ="Paid"
    Expression ="Rmkt.InvoicePaidDate"
    Alias ="SourceCompany"
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="Sch|Grp"
    Expression ="[Schedule] & \"|\" & [UnitGroup]"
    Expression ="Rmkt.BuyerRef"
    Expression ="Rmkt.TargetSale"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Rmkt"
    Expression ="vw_SixKeys.UnitID = Rmkt.UnitRef"
    Flag =1
    LeftTable ="Rmkt"
    RightTable ="Clients"
    Expression ="Rmkt.BuyerRef = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="Rmkt.InvoiceRequestNum"
    Flag =1
    Expression ="vw_SixKeys.AssetID"
    Flag =0
    Expression ="vw_SixKeys.Status"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sch|Grp"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Requested"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Num"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paid"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="4065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sold Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =-8
    Right =1268
    Bottom =973
    Left =-1
    Top =-1
    Right =1198
    Bottom =306
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =614
        Top =83
        Right =944
        Bottom =440
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =296
        Top =76
        Right =568
        Bottom =478
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =1022
        Top =214
        Right =1166
        Bottom =358
        Top =0
        Name ="Clients"
        Name =""
    End
End
