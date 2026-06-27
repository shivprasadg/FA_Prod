Operation =1
Option =0
Begin InputTables
    Name ="ScheduleDocsReportExcelBase"
End
Begin OutputColumns
    Expression ="ScheduleDocsReportExcelBase.Schedule"
    Expression ="ScheduleDocsReportExcelBase.UnitGroup"
    Expression ="ScheduleDocsReportExcelBase.Lessee"
    Expression ="ScheduleDocsReportExcelBase.FAID"
    Alias ="Units"
    Expression ="Count(ScheduleDocsReportExcelBase.Units)"
    Expression ="ScheduleDocsReportExcelBase.UnitDesc"
    Alias ="UnitCost"
    Expression ="Sum(ScheduleDocsReportExcelBase.UnitCost)"
    Alias ="Child Cost"
    Expression ="Sum(ScheduleDocsReportExcelBase.[Child Cost])"
    Alias ="OEC"
    Expression ="Sum(ScheduleDocsReportExcelBase.OEC)"
    Alias ="InvoicePaid"
    Expression ="Max(ScheduleDocsReportExcelBase.InvoicePaid)"
    Expression ="ScheduleDocsReportExcelBase.LeaseTermPrimary"
    Expression ="ScheduleDocsReportExcelBase.LeasePaidAdvArr"
    Expression ="ScheduleDocsReportExcelBase.[Allowed Miles]"
    Expression ="ScheduleDocsReportExcelBase.MileChrg"
    Expression ="ScheduleDocsReportExcelBase.MilesThreshold"
    Expression ="ScheduleDocsReportExcelBase.[MileChrg After]"
    Expression ="ScheduleDocsReportExcelBase.[Term Year]"
    Expression ="ScheduleDocsReportExcelBase.Rate"
    Expression ="ScheduleDocsReportExcelBase.Date"
    Expression ="ScheduleDocsReportExcelBase.[Neg BLCD]"
    Expression ="ScheduleDocsReportExcelBase.[PerDiem Start]"
    Expression ="ScheduleDocsReportExcelBase.BLCD_Projected"
    Expression ="ScheduleDocsReportExcelBase.Assignee"
    Expression ="ScheduleDocsReportExcelBase.Spread"
    Expression ="ScheduleDocsReportExcelBase.Yield"
    Expression ="ScheduleDocsReportExcelBase.Adjustment"
    Expression ="ScheduleDocsReportExcelBase.DeliveryLocation"
    Expression ="ScheduleDocsReportExcelBase.Title"
    Expression ="ScheduleDocsReportExcelBase.Domicle"
    Expression ="ScheduleDocsReportExcelBase.Billing"
    Expression ="ScheduleDocsReportExcelBase.BillingInstructions"
    Expression ="ScheduleDocsReportExcelBase.SalesTax"
    Expression ="ScheduleDocsReportExcelBase.RentReceipt"
    Expression ="ScheduleDocsReportExcelBase.AvalException"
    Expression ="ScheduleDocsReportExcelBase.TaxInstructions"
    Expression ="ScheduleDocsReportExcelBase.IRP"
    Expression ="ScheduleDocsReportExcelBase.SchID"
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
    Expression ="ScheduleDocsReportExcelBase.UnitDesc"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.LeaseTermPrimary"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.LeasePaidAdvArr"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.[Allowed Miles]"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.MileChrg"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.MilesThreshold"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.[MileChrg After]"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.[Term Year]"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Rate"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Date"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.[Neg BLCD]"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.[PerDiem Start]"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.BLCD_Projected"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Assignee"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Spread"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Yield"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Adjustment"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.DeliveryLocation"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Title"
    GroupLevel =0
    Expression ="ScheduleDocsReportExcelBase.Domicle"
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
    Expression ="ScheduleDocsReportExcelBase.SchID"
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
        dbText "Name" ="InvoicePaid"
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
End
Begin
    State =0
    Left =13
    Top =21
    Right =1499
    Bottom =1007
    Left =-1
    Top =-1
    Right =1468
    Bottom =590
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =282
        Top =-21
        Right =426
        Bottom =526
        Top =0
        Name ="ScheduleDocsReportExcelBase"
        Name =""
    End
End
