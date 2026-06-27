Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="RMKT_BuyersOrders"
    Name ="Rmkt"
    Name ="RmktInvoiceRequestTbl"
    Name ="Buyer"
End
Begin OutputColumns
    Expression ="Buyer.BuyerID"
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
    Expression ="Rmkt.SoldPrice"
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
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceRequestNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmodel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1012"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceRequested"
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
        dbText "Name" ="Rmkt.RmktDeposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.rmktdepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer.BuyerName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer.BuyerShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =148
    Top =-9
    Right =1639
    Bottom =1205
    Left =-1
    Top =-1
    Right =1467
    Bottom =900
    Left =604
    Top =0
    ColumnsShown =539
    Begin
        Left =-556
        Top =12
        Right =-412
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =-378
        Top =151
        Right =-234
        Bottom =295
        Top =0
        Name ="RMKT_BuyersOrders"
        Name =""
    End
    Begin
        Left =-172
        Top =12
        Right =-28
        Bottom =156
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =20
        Top =12
        Right =164
        Bottom =156
        Top =0
        Name ="RmktInvoiceRequestTbl"
        Name =""
    End
    Begin
        Left =140
        Top =283
        Right =284
        Bottom =427
        Top =0
        Name ="Buyer"
        Name =""
    End
End
