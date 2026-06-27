Operation =1
Option =0
Where ="(((vw_SixKeys.AcceptedDate) Between [Forms]![LORPTM]![AnalSt] And [Forms]![LORPT"
    "M]![AnalEnd]))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="InvoiceDate"
    Expression ="vw_SixKeys.UnitInvoiceDate"
    Alias ="AcceptedMonth"
    Expression ="Format([AcceptedDate],\"mmm\")"
    Alias ="AcceptedYear"
    Expression ="Format([AcceptedDate],\"yyyy\")"
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
    Expression ="vw_SixKeys.MYear"
    Expression ="vw_SixKeys.Model"
    Expression ="vw_SixKeys.Type"
    Expression ="vw_SixKeys.BuildDate"
    Expression ="vw_SixKeys.ActualDeliveryDate"
    Expression ="vw_SixKeys.AcceptedDate"
    Expression ="vw_SixKeys.InServiceDate"
    Alias ="TermMonths"
    Expression ="SchGrp.LeaseTermPrimary"
    Alias ="LeaseToEnd"
    Expression ="DateAdd(\"m\",[LeaseTermPrimary],[BLCD])"
    Expression ="SchGrp.ExchangeITMonth"
    Alias ="CanExchangeIT"
    Expression ="DateAdd(\"m\",[ExchangeITMonth],[BLCD])"
    Alias ="ActualOffLease"
    Expression ="vw_SixKeys.OffLeaseDate"
    Expression ="SchGrp.BLCD"
    Expression ="vw_SixKeys.UnitCost"
    Alias ="Rent"
    Expression ="Units.UNITRENT"
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
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.PurchaseOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.BuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.EsitmatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ActualDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.AcceptedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MSOReceived"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MSOCopyReceived"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Days2Delivery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.GotMSO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitInvoiceDate"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VendorId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WarrantyCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OtherCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeesTaxesCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DecalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FreightCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AcceptedYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AcceptedMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.PortfolioNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ReportStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.DaysOverdue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.BankId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.OffLeaseDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.PurchaseOrderDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseToEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CanExchangeIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ObcUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TermMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =123
    Top =74
    Right =1502
    Bottom =860
    Left =-1
    Top =-1
    Right =1347
    Bottom =362
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =43
        Top =42
        Right =204
        Bottom =485
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =436
        Top =179
        Right =580
        Bottom =323
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =444
        Top =12
        Right =588
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
