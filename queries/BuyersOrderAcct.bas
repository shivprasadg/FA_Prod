Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
    Name ="RMKT_BuyersOrders"
    Name ="RmktInvoiceRequestTbl"
    Name ="Buyer"
End
Begin OutputColumns
    Expression ="RmktInvoiceRequestTbl.BuyerID"
    Alias ="InvoiceRequestNum"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID"
    Alias ="AssetID"
    Expression ="Units.UnitID"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitVIN"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unitmodel"
    Alias ="Stype"
    Expression ="IIf(Nz([unitSubType],\"\")=\"\",\"\",[UnitSubType]) & IIf([UnitType]=\"Tractor\""
        ",\"\",\" \" & [UnitType])"
    Expression ="Units.unittype"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestDate"
    Expression ="Rmkt.SoldPrice"
    Expression ="RmktInvoiceRequestTbl.InvoiceSentDate"
    Expression ="Buyer.BuyerName"
    Alias ="InvoiceRequested"
    Expression ="0"
    Expression ="Units.unitstatus"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Buyer.BuyerShortName"
    Expression ="Buyer.BuyerID"
    Expression ="Rmkt.RmktDeposit"
    Expression ="Rmkt.rmktdepdt"
    Expression ="Rmkt.pbos"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="RMKT_BuyersOrders"
    Expression ="Units.UnitID = RMKT_BuyersOrders.UnitID"
    Flag =1
    LeftTable ="RMKT_BuyersOrders"
    RightTable ="Rmkt"
    Expression ="RMKT_BuyersOrders.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="RMKT_BuyersOrders"
    RightTable ="RmktInvoiceRequestTbl"
    Expression ="RMKT_BuyersOrders.BuyersOrderID = RmktInvoiceRequestTbl.InvoiceRequestID"
    Flag =1
    LeftTable ="RmktInvoiceRequestTbl"
    RightTable ="Buyer"
    Expression ="RmktInvoiceRequestTbl.BuyerID = Buyer.BuyerID"
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
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktDeposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.rmktdepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.pbos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stype"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceRequested"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceRequestNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer.BuyerShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer.BuyerName"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =286
    Top =215
    Right =1544
    Bottom =861
    Left =-1
    Top =-1
    Right =1234
    Bottom =298
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =754
        Top =129
        Right =981
        Bottom =351
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =488
        Top =216
        Right =714
        Bottom =526
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =261
        Top =93
        Right =404
        Bottom =371
        Top =0
        Name ="RMKT_BuyersOrders"
        Name =""
    End
    Begin
        Left =32
        Top =28
        Right =209
        Bottom =316
        Top =0
        Name ="RmktInvoiceRequestTbl"
        Name =""
    End
    Begin
        Left =893
        Top =53
        Right =1037
        Bottom =197
        Top =0
        Name ="Buyer"
        Name =""
    End
End
