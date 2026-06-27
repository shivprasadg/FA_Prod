Operation =1
Option =0
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Cost_Freight"
    Name ="Cost_Decals"
    Name ="Cost_Refrig"
    Name ="Cost_FeesTax"
    Name ="Cost_OBC"
    Name ="Cost_Warranty"
    Name ="Cost_Other"
    Name ="Cost_OBC_Install"
End
Begin OutputColumns
    Alias ="AcceptedMonth"
    Expression ="Format([AcceptedDate],\"mmm\")"
    Alias ="AcceptedYear"
    Expression ="Format([AcceptedDate],\"yyyy\")"
    Expression ="vw_SixKeys.*"
    Alias ="FreightCost"
    Expression ="CCur(Nz([Freight],0))"
    Alias ="DecalCost"
    Expression ="CCur(Nz([Decals],0))"
    Alias ="RefrigCost"
    Expression ="CCur(Nz([Refrig],0))"
    Alias ="FeesTaxesCost"
    Expression ="CCur(Nz([FeesTax],0))"
    Alias ="OBCMake"
    Expression ="Nz([Cost_OBC].[OBC_Make],\"--\")"
    Alias ="OBCCost"
    Expression ="CCur(Nz([Cost_OBC].[OBC],0))"
    Alias ="OBCInvoice"
    Expression ="Nz(Cost_OBC.[OBCInvoice],\"--\")"
    Alias ="OBCInstaller"
    Expression ="Nz([Cost_OBC_Install].[OBCInstaller],\"--\")"
    Alias ="OBCInstall"
    Expression ="CCur(Nz(Cost_OBC_Install.[OBCInstall],0))"
    Alias ="OBCInstall_Invoice"
    Expression ="Nz(Cost_OBC_Install.[OBCInsInvoice],\"--\")"
    Alias ="WarrantyCost"
    Expression ="CCur(Nz([Warranty],0))"
    Alias ="OtherCost"
    Expression ="CCur(Nz([Other],0))"
    Alias ="InvoiceDate"
    Expression ="vw_SixKeys.UnitInvoiceDate"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Cost_Freight"
    Expression ="vw_SixKeys.AssetID = Cost_Freight.AssetId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Cost_Decals"
    Expression ="vw_SixKeys.AssetID = Cost_Decals.AssetId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Cost_Refrig"
    Expression ="vw_SixKeys.AssetID = Cost_Refrig.AssetId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Cost_FeesTax"
    Expression ="vw_SixKeys.AssetID = Cost_FeesTax.AssetId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Cost_OBC"
    Expression ="vw_SixKeys.AssetID = Cost_OBC.AssetId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Cost_Warranty"
    Expression ="vw_SixKeys.AssetID = Cost_Warranty.AssetId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Cost_Other"
    Expression ="vw_SixKeys.AssetID = Cost_Other.AssetId"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Cost_OBC_Install"
    Expression ="vw_SixKeys.AssetID = Cost_OBC_Install.AssetId"
    Flag =2
End
Begin OrderBy
    Expression ="vw_SixKeys.UnitInvoiceDate"
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
dbMemo "Filter" ="(([AllLeasedUnits_Cost].[clientID]=4971)) AND ([AllLeasedUnits_Cost].[InvoiceDat"
    "e] Is Not Null)"
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
        dbText "Name" ="InvoiceDate"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.DaysOverdue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObcInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstall_Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstaller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ReportStartDate"
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
        dbText "Name" ="UnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =32
    Top =29
    Right =1628
    Bottom =813
    Left =-1
    Top =-1
    Right =1564
    Bottom =328
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
        Left =252
        Top =24
        Right =360
        Bottom =177
        Top =0
        Name ="Cost_Freight"
        Name =""
    End
    Begin
        Left =386
        Top =45
        Right =482
        Bottom =187
        Top =0
        Name ="Cost_Decals"
        Name =""
    End
    Begin
        Left =509
        Top =63
        Right =607
        Bottom =207
        Top =0
        Name ="Cost_Refrig"
        Name =""
    End
    Begin
        Left =625
        Top =79
        Right =727
        Bottom =223
        Top =0
        Name ="Cost_FeesTax"
        Name =""
    End
    Begin
        Left =747
        Top =93
        Right =852
        Bottom =237
        Top =0
        Name ="Cost_OBC"
        Name =""
    End
    Begin
        Left =1005
        Top =138
        Right =1106
        Bottom =282
        Top =0
        Name ="Cost_Warranty"
        Name =""
    End
    Begin
        Left =1121
        Top =164
        Right =1230
        Bottom =308
        Top =0
        Name ="Cost_Other"
        Name =""
    End
    Begin
        Left =875
        Top =114
        Right =987
        Bottom =258
        Top =0
        Name ="Cost_OBC_Install"
        Name =""
    End
End
