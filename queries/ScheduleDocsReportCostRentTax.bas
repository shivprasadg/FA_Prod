Operation =1
Option =0
Begin InputTables
    Name ="ScheduleDocsReportExcelBase"
    Name ="Units"
End
Begin OutputColumns
    Expression ="ScheduleDocsReportExcelBase.Schedule"
    Expression ="ScheduleDocsReportExcelBase.UnitGroup"
    Expression ="ScheduleDocsReportExcelBase.Lessee"
    Expression ="ScheduleDocsReportExcelBase.FAID"
    Alias ="AssetID"
    Expression ="ScheduleDocsReportExcelBase.Units"
    Expression ="ScheduleDocsReportExcelBase.UnitNum"
    Expression ="ScheduleDocsReportExcelBase.VIN"
    Expression ="ScheduleDocsReportExcelBase.UnitDesc"
    Alias ="UnitCost"
    Expression ="Sum(ScheduleDocsReportExcelBase.UnitCost)"
    Alias ="Child Cost"
    Expression ="Sum(ScheduleDocsReportExcelBase.[Child Cost])"
    Alias ="OEC"
    Expression ="Sum(ScheduleDocsReportExcelBase.OEC)"
    Alias ="VendorCost"
    Expression ="Sum(ScheduleDocsReportExcelBase.VendorCost)"
    Alias ="SchRent"
    Expression ="Sum(ScheduleDocsReportExcelBase.SchRent)"
    Alias ="SchRentLRF"
    Expression ="Sum([ScheduleDocsReportExcelBase].[schrent])/Sum([ScheduleDocsReportExcelBase].["
        "OEC])"
    Alias ="ExtdRent"
    Expression ="Sum(ScheduleDocsReportExcelBase.ExtdRent)"
    Alias ="ExtRentLRF"
    Expression ="Sum([ScheduleDocsReportExcelBase].[Extdrent])/Sum([ScheduleDocsReportExcelBase]."
        "[OEC])"
    Alias ="InvoicePaid"
    Expression ="Max(ScheduleDocsReportExcelBase.InvoicePaid)"
    Expression ="ScheduleDocsReportExcelBase.LeaseTermPrimary"
    Expression ="ScheduleDocsReportExcelBase.LeasePaidAdvArr"
    Expression ="ScheduleDocsReportExcelBase.[Neg BLCD]"
    Expression ="ScheduleDocsReportExcelBase.[PerDiem Start]"
    Expression ="ScheduleDocsReportExcelBase.BLCD_Projected"
    Expression ="ScheduleDocsReportExcelBase.Billing"
    Expression ="ScheduleDocsReportExcelBase.BillingInstructions"
    Expression ="ScheduleDocsReportExcelBase.SalesTax"
    Expression ="ScheduleDocsReportExcelBase.RentReceipt"
    Expression ="ScheduleDocsReportExcelBase.AvalException"
    Expression ="ScheduleDocsReportExcelBase.TaxInstructions"
    Expression ="ScheduleDocsReportExcelBase.IRP"
    Expression ="ScheduleDocsReportExcelBase.FATitleLogNum"
    Expression ="ScheduleDocsReportExcelBase.SchID"
    Alias ="EstDelv"
    Expression ="Units.UnitEstDelDate"
    Alias ="ActualDelv"
    Expression ="Units.UnitActualDelDate"
    Alias ="PDRent"
    Expression ="Units.UNITPDRENT"
    Alias ="PDStart"
    Expression ="Units.UnitPDStartDate"
End
Begin Joins
    LeftTable ="ScheduleDocsReportExcelBase"
    RightTable ="Units"
    Expression ="ScheduleDocsReportExcelBase.Units = Units.UnitID"
    Flag =1
End
Begin Groups
    Expression ="ScheduleDocsReportExcelBase.Schedule"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.UnitGroup"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Lessee"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.FAID"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Units"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.UnitNum"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.VIN"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.UnitDesc"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.LeaseTermPrimary"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.LeasePaidAdvArr"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.[Neg BLCD]"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.[PerDiem Start]"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.BLCD_Projected"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Billing"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.BillingInstructions"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.SalesTax"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.RentReceipt"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.AvalException"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.TaxInstructions"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.IRP"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.FATitleLogNum"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.SchID"
    GroupLevel =0
    Expression ="Units.UnitEstDelDate"
    GroupLevel =0
    Expression ="Units.UnitActualDelDate"
    GroupLevel =0
    Expression ="Units.UNITPDRENT"
    GroupLevel =0
    Expression ="Units.UnitPDStartDate"
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
        dbText "Name" ="ScheduleDocsReportExcelBase.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.[Term Year]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Rate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.MileChrg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.MilesThreshold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Adjustment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.DeliveryLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.[Allowed Miles]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.[PerDiem Start]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.[MileChrg After]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.[Neg BLCD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Spread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Yield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.AvalException"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfChild Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.SalesTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Domicle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.[Child Cost]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Child Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.RentReceipt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.TaxInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.BillingInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.IRP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.Billing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchRentLRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.ExtdRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfExtdRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtRentLRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtdRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.FATitleLogNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EstDelv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualDelv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDocsReportExcelBase.VIN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =302
    Top =53
    Right =1788
    Bottom =1039
    Left =-1
    Top =-1
    Right =1462
    Bottom =437
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =259
        Top =-9
        Right =403
        Bottom =413
        Top =0
        Name ="ScheduleDocsReportExcelBase"
        Name =""
    End
    Begin
        Left =501
        Top =177
        Right =645
        Bottom =461
        Top =0
        Name ="Units"
        Name =""
    End
End
