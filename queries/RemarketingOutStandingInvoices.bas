Operation =1
Option =0
Where ="(((RMKT.InvoiceRequestNum) Is Not Null) AND ((RMKT.InvoicePaidDate) Is Null))"
Begin InputTables
    Name ="RMKT"
    Name ="Buyers"
End
Begin OutputColumns
    Expression ="Buyers.ClientShNm"
    Expression ="RMKT.InvoiceRequestNum"
    Expression ="RMKT.InvoiceRequestDate"
    Expression ="RMKT.InvoiceSentDate"
    Expression ="RMKT.InvoicePaidDate"
    Expression ="RMKT.UnitRef"
    Expression ="RMKT.BuyerRef"
End
Begin Joins
    LeftTable ="RMKT"
    RightTable ="Buyers"
    Expression ="RMKT.BuyerRef = Buyers.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="Buyers.ClientShNm"
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
        dbText "Name" ="RMKT.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKT.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKT.InvoiceRequestNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKT.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKT.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Buyers.ClientShNm"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKT.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =60
    Top =23
    Right =1498
    Bottom =951
    Left =-1
    Top =-1
    Right =1406
    Bottom =610
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =334
        Bottom =403
        Top =0
        Name ="RMKT"
        Name =""
    End
    Begin
        Left =392
        Top =138
        Right =536
        Bottom =282
        Top =0
        Name ="Buyers"
        Name =""
    End
End
