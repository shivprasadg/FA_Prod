Operation =1
Option =0
Begin InputTables
    Name ="RMKT_BuyersOrders"
    Alias ="RBO"
    Name ="RmktInvoiceRequestTbl"
    Alias ="tReq"
    Name ="Units"
    Alias ="U"
    Name ="Rmkt"
    Alias ="R"
End
Begin OutputColumns
    Expression ="RBO.BuyersOrderID"
    Expression ="RBO.UnitID"
    Expression ="U.UnitUnitNum"
    Expression ="U.UnitVIN"
    Expression ="tReq.InvoiceDate"
    Expression ="tReq.InvoiceSentDate"
    Expression ="tReq.InvoicePaidDate"
    Expression ="R.TargetSale"
    Alias ="Days"
    Expression ="DateDiff(\"d\",tReq.InvoiceSentDate,Nz(tReq.InvoicePaidDate,Date()))"
    Expression ="RBO.SalePrice"
    Expression ="tReq.FundingType"
    Alias ="UnitDesc"
    Expression ="CStr(UnitYr) & \"  \" & U.UnitMake & \"  \" & U.UnitModel & \"  \" & U.UnitSubTy"
        "pe & \"  \" & U.UnitType+' '+U.Axle"
    Expression ="tReq.SP"
    Alias ="STS"
    Expression ="U.UnitStatus"
    Expression ="tReq.BuyerID"
    Expression ="RBO.WarrantyPrice"
End
Begin Joins
    LeftTable ="RBO"
    RightTable ="tReq"
    Expression ="RBO.BuyersOrderID = tReq.InvoiceRequestID"
    Flag =1
    LeftTable ="RBO"
    RightTable ="U"
    Expression ="RBO.UnitID = U.UnitID"
    Flag =1
    LeftTable ="U"
    RightTable ="R"
    Expression ="U.UnitID = R.UnitRef"
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
        dbText "Name" ="RBO.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tReq.InvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tReq.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tReq.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Days"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RBO.SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tReq.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tReq.FundingType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tReq.SP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="STS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RBO.BuyersOrderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RBO.WarrantyPrice"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =146
    Top =60
    Right =1630
    Bottom =1364
    Left =-1
    Top =-1
    Right =1460
    Bottom =257
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =227
        Top =67
        Right =371
        Bottom =211
        Top =0
        Name ="RBO"
        Name =""
    End
    Begin
        Left =45
        Top =22
        Right =189
        Bottom =166
        Top =0
        Name ="tReq"
        Name =""
    End
    Begin
        Left =409
        Top =69
        Right =553
        Bottom =213
        Top =0
        Name ="U"
        Name =""
    End
    Begin
        Left =591
        Top =73
        Right =735
        Bottom =217
        Top =0
        Name ="R"
        Name =""
    End
End
