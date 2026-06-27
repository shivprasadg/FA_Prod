Operation =5
Option =8
Where ="(((Rmkt_WrongStatus.Inspected) Is Null) AND ((Rmkt.BuyerRef) Is Null) AND ((Rmkt"
    ".BuyOutCost) Is Null) AND ((Rmkt.InvoicePaidDate) Is Null) AND ((Rmkt.[SUF Sent]"
    ") Is Null))"
Begin InputTables
    Name ="Rmkt"
    Name ="Rmkt_WrongStatus"
End
Begin OutputColumns
    Expression ="Rmkt.*"
    Expression ="Rmkt_WrongStatus.Inspected"
    Expression ="Rmkt.BuyerRef"
    Expression ="Rmkt.BuyOutCost"
    Expression ="Rmkt.InvoicePaidDate"
    Expression ="Rmkt.[SUF Sent]"
End
Begin Joins
    LeftTable ="Rmkt"
    RightTable ="Rmkt_WrongStatus"
    Expression ="Rmkt.UnitRef = Rmkt_WrongStatus.AssetID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Rmkt_WrongStatus.Inspected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmktId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SUF Sent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SUF Returned"
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
        dbText "Name" ="Rmkt.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyerRef"
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
        dbText "Name" ="Rmkt.Insp by2"
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
        dbText "Name" ="Rmkt.IFBankAcctNumber"
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
    Begin
        dbText "Name" ="Rmkt.Commision"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.[SUF Sent]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1676
    Bottom =872
    Left =-1
    Top =-1
    Right =1644
    Bottom =498
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =202
        Top =71
        Right =366
        Bottom =425
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =433
        Top =74
        Right =633
        Bottom =550
        Top =0
        Name ="Rmkt_WrongStatus"
        Name =""
    End
End
