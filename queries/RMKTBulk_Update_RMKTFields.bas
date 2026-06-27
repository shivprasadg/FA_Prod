Operation =1
Option =0
Begin InputTables
    Name ="RMKTBulk_Update_WorkTable"
    Name ="Rmkt"
End
Begin OutputColumns
    Alias ="AssetID"
    Expression ="Rmkt.UnitRef"
    Expression ="Rmkt.[SUF Sent]"
    Expression ="Rmkt.[SUF Returned]"
    Expression ="Rmkt.ParkDT"
    Expression ="Rmkt.BuyOutDate"
    Expression ="Rmkt.BuyOutCost"
    Expression ="Rmkt.BuyerRef"
    Expression ="Rmkt.SoldPrice"
    Expression ="Rmkt.RmktSPRAdj"
    Expression ="Rmkt.TargetSale"
    Expression ="Rmkt.SoldPrice"
    Expression ="Rmkt.InvoiceSentDate"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Rmkt.RmktBOSSent"
    Expression ="Rmkt.RmKTBOSRec"
    Expression ="Rmkt.RmktTitleRec"
    Expression ="Rmkt.RmktTitleSent"
    Expression ="Rmkt.BankNotified"
    Expression ="Rmkt.Inspected"
    Expression ="Rmkt.InspectedBy"
    Expression ="Rmkt.InspChg"
    Expression ="Rmkt.InspResults"
    Expression ="Rmkt.RmtTitletoFA"
    Expression ="Rmkt.TitleOut"
    Expression ="Rmkt.FATitleBack"
    Expression ="Rmkt.MiParked"
    Expression ="Rmkt.HrParked"
    Expression ="Rmkt.RmkSm"
    Expression ="Rmkt.RmkLOSupport"
    Expression ="Rmkt.Inspected2"
    Expression ="Rmkt.[Insp by2]"
    Expression ="Rmkt.InspChg2"
    Expression ="Rmkt.InspResults2"
    Expression ="Rmkt.InspInv1"
    Expression ="Rmkt.InspInv2"
    Expression ="Rmkt.FreeStrDays"
    Expression ="Rmkt.FreeStrexp"
    Expression ="Rmkt.AdjBookValue"
    Expression ="Rmkt.RmktAdjNote"
    Expression ="Rmkt.InvoiceRequested"
    Expression ="Rmkt.InvoiceRequestDate"
    Expression ="Rmkt.PBOS"
    Expression ="Rmkt.RmktDeposit"
    Expression ="Rmkt.RmktDepdt"
    Expression ="Rmkt.WhSValue"
    Expression ="Rmkt.Custbookgoal"
    Expression ="Rmkt.FABrkComm"
    Expression ="Rmkt.upsize_ts"
    Expression ="Rmkt.InspMileage1"
    Expression ="Rmkt.InspMileage2"
    Expression ="Rmkt.PRBOS"
    Expression ="Rmkt.Discount"
    Expression ="Rmkt.DiscountPercentage"
    Expression ="Rmkt.Shipping"
    Expression ="Rmkt.NoScrapSalvageSale"
    Expression ="Rmkt.IFBankID"
    Expression ="Rmkt.IFLoanNumber"
    Expression ="Rmkt.IFDate"
    Expression ="Rmkt.IFAmount"
    Expression ="Rmkt.IFPercentage"
    Expression ="Rmkt.FAAmount"
    Expression ="Rmkt.Commision"
End
Begin Joins
    LeftTable ="RMKTBulk_Update_WorkTable"
    RightTable ="Rmkt"
    Expression ="RMKTBulk_Update_WorkTable.UnitID = Rmkt.UnitRef"
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
        dbText "Name" ="Rmkt.Commision"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.[SUF Sent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.[SUF Returned]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.ParkDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSPRAdj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktBOSSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmKTBOSRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktTitleRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktTitleSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BankNotified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Inspected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspectedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspChg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspResults"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmtTitletoFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TitleOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FATitleBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.MiParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.HrParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmkSm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmkLOSupport"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Inspected2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.[Insp by2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspChg2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspResults2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspInv1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspInv2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FreeStrDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FreeStrexp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.AdjBookValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktAdjNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PInvReq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.AInvREq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PBOS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktDeposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktDepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.WhSValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Custbookgoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FABrkComm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.upsize_ts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspMileage1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InspMileage2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PRBOS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.DiscountPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Shipping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.NoScrapSalvageSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FAAmount"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =89
    Top =24
    Right =1706
    Bottom =865
    Left =-1
    Top =-1
    Right =1585
    Bottom =302
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =111
        Top =37
        Right =255
        Bottom =181
        Top =0
        Name ="RMKTBulk_Update_WorkTable"
        Name =""
    End
    Begin
        Left =336
        Top =49
        Right =507
        Bottom =320
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
