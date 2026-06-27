Operation =1
Option =0
Begin InputTables
    Name ="DrawDowns"
    Name ="Units"
    Name ="Banks"
    Name ="VW_Clients_ML_FAID_SCH_Groups_count"
End
Begin OutputColumns
    Expression ="DrawDowns.DrawDownId"
    Alias ="Bank"
    Expression ="Banks.BankShortName"
    Expression ="DrawDowns.DrawNumber"
    Expression ="DrawDowns.DrawBankAccountNumber"
    Expression ="DrawDowns.DrawRequestNumber"
    Alias ="AssetID"
    Expression ="Units.UnitID"
    Alias ="VIN"
    Expression ="Units.UnitVIN"
    Alias ="UnitNumber"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitStatus"
    Alias ="ActualDelivery"
    Expression ="Units.UnitActualDelDate"
    Alias ="InvoiceNum"
    Expression ="Units.UnitInvoiceNum"
    Alias ="InvoiceDate"
    Expression ="Units.UnitInvoiceDate"
    Alias ="VendorPaid"
    Expression ="Units.UnitVendorPytDate"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.ClientShortName"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.MLNo"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.FAID"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.Schedule"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.UnitGroup"
    Expression ="DrawDowns.ClientID"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.SchID"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.SGrpID"
    Alias ="VendorID"
    Expression ="Units.VendorREF"
End
Begin Joins
    LeftTable ="DrawDowns"
    RightTable ="Units"
    Expression ="DrawDowns.DrawDownId = Units.DrawDownId"
    Flag =1
    LeftTable ="DrawDowns"
    RightTable ="Banks"
    Expression ="DrawDowns.DrawBankId = Banks.BankId"
    Flag =1
    LeftTable ="Units"
    RightTable ="VW_Clients_ML_FAID_SCH_Groups_count"
    Expression ="Units.GroupREF = VW_Clients_ML_FAID_SCH_Groups_count.SGrpID"
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
dbMemo "Filter" ="([DrawMatch2Units].[InvoiceNum]=\"917574\")"
Begin
    Begin
        dbText "Name" ="DrawDowns.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualDelivery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawNumber"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawBankAccountNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawRequestNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.VendorId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVendorPytDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.VendorId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =37
    Right =1571
    Bottom =889
    Left =-1
    Top =-1
    Right =1487
    Bottom =466
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =96
        Right =219
        Bottom =385
        Top =0
        Name ="DrawDowns"
        Name =""
    End
    Begin
        Left =407
        Top =54
        Right =609
        Bottom =430
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =249
        Top =558
        Right =393
        Bottom =702
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =686
        Top =166
        Right =853
        Bottom =411
        Top =0
        Name ="VW_Clients_ML_FAID_SCH_Groups_count"
        Name =""
    End
End
