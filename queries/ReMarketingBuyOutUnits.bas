Operation =1
Option =8
Begin InputTables
    Name ="ReMarketingBuyOut_WorkTable"
    Name ="RemarketingUnitsBySchedule"
End
Begin OutputColumns
    Expression ="RemarketingUnitsBySchedule.UnitID"
    Alias ="UnitNum"
    Expression ="RemarketingUnitsBySchedule.Unitunitnum"
    Expression ="RemarketingUnitsBySchedule.Unitvin"
    Expression ="RemarketingUnitsBySchedule.SchID"
    Expression ="RemarketingUnitsBySchedule.SchNo"
    Expression ="RemarketingUnitsBySchedule.UnitStatus"
    Expression ="RemarketingUnitsBySchedule.BuyOutDate"
    Expression ="RemarketingUnitsBySchedule.BuyOutCost"
    Expression ="RemarketingUnitsBySchedule.IFBankID"
    Expression ="RemarketingUnitsBySchedule.IFBank"
    Expression ="RemarketingUnitsBySchedule.IFLoanNumber"
    Expression ="RemarketingUnitsBySchedule.IFDate"
    Expression ="RemarketingUnitsBySchedule.IFAmount"
    Expression ="RemarketingUnitsBySchedule.IFPaidDate"
    Expression ="RemarketingUnitsBySchedule.InvoiceSentDate"
    Expression ="RemarketingUnitsBySchedule.InvoicePaidDate"
    Expression ="RemarketingUnitsBySchedule.SoldPrice"
    Expression ="RemarketingUnitsBySchedule.BuyDateEffective"
End
Begin Joins
    LeftTable ="ReMarketingBuyOut_WorkTable"
    RightTable ="RemarketingUnitsBySchedule"
    Expression ="ReMarketingBuyOut_WorkTable.UnitID = RemarketingUnitsBySchedule.UnitID"
    Flag =1
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
        dbText "Name" ="RemarketingUnitsBySchedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.Unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.IFPaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.IFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.IFAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.IFBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.IFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingUnitsBySchedule.BuyDateEffective"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =358
    Top =144
    Right =1901
    Bottom =638
    Left =-1
    Top =-1
    Right =1519
    Bottom =164
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =24
        Top =36
        Right =168
        Bottom =180
        Top =0
        Name ="ReMarketingBuyOut_WorkTable"
        Name =""
    End
    Begin
        Left =219
        Top =12
        Right =407
        Bottom =254
        Top =0
        Name ="RemarketingUnitsBySchedule"
        Name =""
    End
End
