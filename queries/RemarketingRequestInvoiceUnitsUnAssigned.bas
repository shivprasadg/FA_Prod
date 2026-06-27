Operation =1
Option =0
Where ="(((vw_SixKeys.Status)=\"R\" Or (vw_SixKeys.Status)=\"PS\" Or (vw_SixKeys.Status)"
    "=\"TR\" Or (vw_SixKeys.Status)=\"P\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Rmkt"
    Name ="RMKT_BuyersOrders"
    Name ="RmktInvoiceRequestTbl"
End
Begin OutputColumns
    Expression ="vw_SixKeys.AssetID"
    Alias ="Unit#"
    Expression ="vw_SixKeys.Unitnum"
    Expression ="vw_SixKeys.VIN"
    Alias ="Description"
    Expression ="[MYear] & \" \" & [Make] & \" \" & [Model] & IIf(Nz([SubType],\"\")=\"\",\"\",\""
        " \" & [SubType]) & IIf([Type]=\"Tractor\",\"\",\" \" & [Type])"
    Expression ="vw_SixKeys.Status"
    Expression ="Rmkt.TargetSale"
    Alias ="SourceCompany"
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="Sch|Grp"
    Expression ="[Schedule] & \"|\" & [UnitGroup]"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID"
    Expression ="RmktInvoiceRequestTbl.InvoiceSentDate"
    Expression ="RmktInvoiceRequestTbl.InvoicePaidDate"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestDate"
    Expression ="RmktInvoiceRequestTbl.BuyerID"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Rmkt"
    Expression ="vw_SixKeys.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="RMKT_BuyersOrders"
    Expression ="vw_SixKeys.UnitID = RMKT_BuyersOrders.UnitID"
    Flag =2
    LeftTable ="RMKT_BuyersOrders"
    RightTable ="RmktInvoiceRequestTbl"
    Expression ="RMKT_BuyersOrders.BuyersOrderID = RmktInvoiceRequestTbl.InvoiceRequestID"
    Flag =2
End
Begin OrderBy
    Expression ="vw_SixKeys.AssetID"
    Flag =0
    Expression ="vw_SixKeys.Status"
    Flag =0
    Expression ="vw_SixKeys.ClientShortName"
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
        dbText "Name" ="SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sch|Grp"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =27
    Top =9
    Right =1257
    Bottom =990
    Left =-1
    Top =-1
    Right =1206
    Bottom =442
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =279
        Top =29
        Right =453
        Bottom =402
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =512
        Top =112
        Right =656
        Bottom =302
        Top =0
        Name ="RMKT_BuyersOrders"
        Name =""
    End
    Begin
        Left =704
        Top =167
        Right =848
        Bottom =433
        Top =0
        Name ="RmktInvoiceRequestTbl"
        Name =""
    End
End
