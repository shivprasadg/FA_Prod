Operation =1
Option =0
Begin InputTables
    Name ="RmktInvoiceRequestTbl"
    Name ="Rmkt"
    Name ="Units"
    Name ="Buyer"
End
Begin OutputColumns
    Expression ="RmktInvoiceRequestTbl.BuyerID"
    Alias ="Buyer"
    Expression ="Buyer.BuyerShortName"
    Alias ="Req"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID"
    Alias ="InvNum"
    Expression ="RmktInvoiceRequestTbl.InvoiceNumber"
    Alias ="Sent"
    Expression ="Format([Rmkt].[InvoiceSentDate],\"mm/dd/yy\")"
    Alias ="AssetID"
    Expression ="Units.UnitID"
    Alias ="UnitNum"
    Expression ="Units.UnitunitNum"
    Alias ="Vin"
    Expression ="Units.UnitVin"
    Alias ="UnitDesc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType]"
    Alias ="Paid"
    Expression ="Format([rmkt].[InvoicePaidDate],\"mm/dd/yy\")"
End
Begin Joins
    LeftTable ="RmktInvoiceRequestTbl"
    RightTable ="Rmkt"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID = Rmkt.InvoiceRequestNum"
    Flag =1
    LeftTable ="Rmkt"
    RightTable ="Units"
    Expression ="Rmkt.UnitRef = Units.UnitID"
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
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.BuyerID"
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyer"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3975"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Req"
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
    Left =39
    Top =5
    Right =1252
    Bottom =986
    Left =-1
    Top =-1
    Right =1189
    Bottom =578
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =37
        Top =59
        Right =231
        Bottom =373
        Top =0
        Name ="RmktInvoiceRequestTbl"
        Name =""
    End
    Begin
        Left =270
        Top =176
        Right =516
        Bottom =572
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =557
        Top =249
        Right =701
        Bottom =393
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =692
        Top =10
        Right =836
        Bottom =154
        Top =0
        Name ="Buyer"
        Name =""
    End
End
