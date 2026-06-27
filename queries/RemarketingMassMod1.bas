Operation =1
Option =0
Where ="(((Units.UnitStatus)<>\"B\"))"
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Rmkt"
    Name ="RMKT_MMod_Work"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Alias ="UnitNum"
    Expression ="Units.Unitunitnum"
    Expression ="Units.Unitvin"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="Units.UnitStatus"
    Expression ="Rmkt.*"
    Expression ="Units.unitoffleasedt"
    Expression ="Units.LocationParkedID"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="RMKT_MMod_Work"
    Expression ="Units.UnitID = RMKT_MMod_Work.UnitID"
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
        dbText "Name" ="Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TitleOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmkSm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmkLOSupport"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PBOS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFLoanNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SUF Sent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SellerInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.IFPaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FATitleBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FAAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktAdjNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SUF Returned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSPRAdj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmKTBOSRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.HrParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.AdjBookValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Commision"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktBOSSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktDeposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FreeStrexp"
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
        dbText "Name" ="Rmkt.MiParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FreeStrDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.ParkDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktTitleRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Shipping"
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
        dbText "Name" ="Rmkt.DiscountPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FMVValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.gainloss"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceRequestNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.MileageProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.ParkDateProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BankShareNetProceeds"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.Wholesale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceRequested"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyDateEffective"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutCostNet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.DrawDownID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LocationPArkedID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =124
    Top =153
    Right =1689
    Bottom =1087
    Left =-1
    Top =-1
    Right =1541
    Bottom =329
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =377
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =801
        Bottom =364
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =87
        Top =195
        Right =231
        Bottom =339
        Top =0
        Name ="RMKT_MMod_Work"
        Name =""
    End
End
