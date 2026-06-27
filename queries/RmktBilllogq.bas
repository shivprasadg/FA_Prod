Operation =1
Option =0
Where ="(((Clients.clientType)=\"remarketing\"))"
Begin InputTables
    Name ="Clients"
    Name ="RmktBills"
    Name ="Units"
    Name ="Rmkt"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.ClientShNm"
    Expression ="RmktBills.[Invoice No]"
    Expression ="RmktBills.RevNo"
    Expression ="RmktBills.Billdt"
    Expression ="RmktBills.RBillDueDt"
    Expression ="RmktBills.BillSent"
    Expression ="Rmkt.InvoiceRequestDate"
    Alias ="CountOfID"
    Expression ="Count(RmktBills.ID)"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Units.unitstatus"
    Alias ="BillSt"
    Expression ="IIf([unitstatus]=\"s\",3,IIf([unitstatus]=\"ks\",4,IIf(IsNull([billsent]),1,2)))"
    Alias ="SumOfSoldPrice"
    Expression ="Sum(Rmkt.SoldPrice)"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="RmktBills"
    Expression ="Clients.clientID = RmktBills.BuyerRef"
    Flag =1
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =3
    LeftTable ="RmktBills"
    RightTable ="Rmkt"
    Expression ="RmktBills.BuyerRef = Rmkt.BuyerRef"
    Flag =2
    LeftTable ="RmktBills"
    RightTable ="Rmkt"
    Expression ="RmktBills.ReqDt = Rmkt.InvoiceRequestDate"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
End
Begin Groups
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="RmktBills.[Invoice No]"
    GroupLevel =0
    Expression ="RmktBills.RevNo"
    GroupLevel =0
    Expression ="RmktBills.Billdt"
    GroupLevel =0
    Expression ="RmktBills.RBillDueDt"
    GroupLevel =0
    Expression ="RmktBills.BillSent"
    GroupLevel =0
    Expression ="Rmkt.InvoiceRequestDate"
    GroupLevel =0
    Expression ="Rmkt.InvoicePaidDate"
    GroupLevel =0
    Expression ="Units.unitstatus"
    GroupLevel =0
    Expression ="IIf([unitstatus]=\"s\",3,IIf([unitstatus]=\"ks\",4,IIf(IsNull([billsent]),1,2)))"
    GroupLevel =0
    Expression ="RmktBills.ID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4245"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.[Invoice No]"
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.RevNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="RmktBills.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.RBillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktBills.BillSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.ainvreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfRmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfID"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =224
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =227
        Top =8
        Right =371
        Bottom =152
        Top =0
        Name ="RmktBills"
        Name =""
    End
    Begin
        Left =671
        Top =11
        Right =815
        Bottom =155
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =474
        Top =10
        Right =618
        Bottom =154
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
