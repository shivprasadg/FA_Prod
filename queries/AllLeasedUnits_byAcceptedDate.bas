Operation =1
Option =0
Where ="(((vw_SixKeys.AcceptedDate) Between [Forms]![LORPTM]![AnalSt] And [Forms]![LORPT"
    "M]![AnalEnd]))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="SchGrp"
    Name ="ClientGroups"
End
Begin OutputColumns
    Expression ="vw_SixKeys.AcceptedDate"
    Alias ="AcceptedMonth"
    Expression ="Format([AcceptedDate],\"mmm\")"
    Alias ="AcceptedYear"
    Expression ="Format([AcceptedDate],\"yyyy\")"
    Alias ="ClientGroup"
    Expression ="ClientGroups.ClientGroupName"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.FAID"
    Expression ="vw_SixKeys.MLNo"
    Expression ="vw_SixKeys.MLDate"
    Expression ="vw_SixKeys.BankShortName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.PurchaseOrder"
    Expression ="vw_SixKeys.PurchaseOrderDate"
    Expression ="vw_SixKeys.UnitID"
    Expression ="vw_SixKeys.Status"
    Expression ="vw_SixKeys.Unitnum"
    Alias ="UnitDesc"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Expression ="vw_SixKeys.BuildDate"
    Alias ="InvoiceDate"
    Expression ="vw_SixKeys.UnitInvoiceDate"
    Expression ="vw_SixKeys.ActualDeliveryDate"
    Expression ="vw_SixKeys.InServiceDate"
    Alias ="TermMonths"
    Expression ="SchGrp.LeaseTermPrimary"
    Alias ="LeaseToEnd"
    Expression ="DateAdd(\"m\",[SchGrp].[LeaseTermPrimary],[SchGrp].[BLCD])"
    Expression ="SchGrp.ExchangeITMonth"
    Alias ="CanExchangeIT"
    Expression ="DateAdd(\"m\",[ExchangeITMonth],SchGrp.[BLCD])"
    Alias ="ActualOffLease"
    Expression ="vw_SixKeys.OffLeaseDate"
    Expression ="SchGrp.BLCD"
    Expression ="vw_SixKeys.UnitCost"
    Alias ="SyndRent"
    Expression ="Units.UNITRENT"
    Alias ="Rent"
    Expression ="Units.UNITSRENT"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="ClientGroups"
    Expression ="vw_SixKeys.ClientGroupId = ClientGroups.ClientGroupId"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.AcceptedDate"
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
        dbText "Name" ="AcceptedYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AcceptedMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TermMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseToEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CanExchangeIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AcceptedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.PurchaseOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.PurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ActualDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescShortVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientGroups.ClientGroupName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =126
    Top =60
    Right =1897
    Bottom =846
    Left =-1
    Top =-1
    Right =1753
    Bottom =277
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
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="ClientGroups"
        Name =""
    End
End
