Operation =1
Option =0
Begin InputTables
    Name ="RmktInvoiceRequestTbl"
    Name ="Rmkt"
    Name ="Clients"
End
Begin OutputColumns
    Alias ="Buyer"
    Expression ="Clients.ClientShNm"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestDate"
    Alias ="Status"
    Expression ="Switch(Not IsDate([Rmkt].[InvoiceRequestDate]),\"Open\",IsDate([Rmkt].[InvoiceRe"
        "questDate]) And Not IsDate([Rmkt].[InvoicePaidDate]) And Not IsDate([Rmkt].[Invo"
        "iceSentDate]),\"Request\",Not IsDate([Rmkt].[InvoicePaidDate]) And IsDate([Rmkt]"
        ".[InvoiceSentDate]),\"Sent\",IsDate([Rmkt].[InvoicePaidDate]),\"Paid\")"
    Alias ="Units"
    Expression ="Count(Rmkt.UnitRef)"
    Expression ="RmktInvoiceRequestTbl.BuyerID"
    Expression ="Rmkt.InvoiceSentDate"
    Expression ="Rmkt.InvoicePaidDate"
    Alias ="UnitRequestDate"
    Expression ="Rmkt.InvoiceRequestDate"
    Alias ="Sent"
    Expression ="RmktInvoiceRequestTbl.InvoiceSentDate"
    Alias ="Paid"
    Expression ="RmktInvoiceRequestTbl.InvoicePaidDate"
End
Begin Joins
    LeftTable ="RmktInvoiceRequestTbl"
    RightTable ="Rmkt"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID = Rmkt.InvoiceRequestNum"
    Flag =2
    LeftTable ="RmktInvoiceRequestTbl"
    RightTable ="Clients"
    Expression ="RmktInvoiceRequestTbl.BuyerID = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID"
    Flag =1
End
Begin Groups
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID"
    GroupLevel =0
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestDate"
    GroupLevel =0
    Expression ="RmktInvoiceRequestTbl.BuyerID"
    GroupLevel =0
    Expression ="Rmkt.InvoiceSentDate"
    GroupLevel =0
    Expression ="Rmkt.InvoicePaidDate"
    GroupLevel =0
    Expression ="Rmkt.InvoiceRequestDate"
    GroupLevel =0
    Expression ="RmktInvoiceRequestTbl.InvoiceSentDate"
    GroupLevel =0
    Expression ="RmktInvoiceRequestTbl.InvoicePaidDate"
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
        dbText "Name" ="RmktInvoiceRequestTbl.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paid"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-240
    Top =56
    Right =1154
    Bottom =931
    Left =-1
    Top =-1
    Right =1362
    Bottom =489
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =97
        Top =43
        Right =433
        Bottom =315
        Top =0
        Name ="RmktInvoiceRequestTbl"
        Name =""
    End
    Begin
        Left =515
        Top =73
        Right =800
        Bottom =561
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =887
        Top =263
        Right =1031
        Bottom =407
        Top =0
        Name ="Clients"
        Name =""
    End
End
