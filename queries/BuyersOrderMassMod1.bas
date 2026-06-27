dbMemo "SQL" ="SELECT Units.UnitID, Units.Unitunitnum AS UnitNum, Units.Unitvin, Schedule.SchID"
    ", Schedule.SchNo, Units.UnitStatus, Last, Rmkt.*, Units.unitoffleasedt, Units.Lo"
    "cationParkedID, Units.clientGroupId, SchGrp.AssigneeBankID\015\012FROM (Schedule"
    " INNER JOIN (((SchGrp INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF) LEFT J"
    "OIN Rmkt ON Units.UnitID = Rmkt.UnitRef) LEFT JOIN (SELECT\015\012              "
    "      AssetID,\015\012                    LastDirection AS Last\015\012         "
    "       FROM\015\012                    RMKTUnitTitleLog_Work\015\012            "
    ")  AS LT ON Units.UnitID = LT.AssetID) ON Schedule.SchID = SchGrp.SGrpSchID) INN"
    "ER JOIN BuyersOrder_MMod_Work ON Units.UnitID = BuyersOrder_MMod_Work.UnitID\015"
    "\012WHERE (((Units.UnitStatus) <> \"B\"));\015\012"
dbMemo "Connect" =""
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
    Begin
        dbText "Name" ="Rmkt.PayOffRequested"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.SufPDFLink"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.FAPortfolio"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PayOffReceived"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.clientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.AssigneeBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.TargetSaleDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last"
        dbLong "AggregateType" ="-1"
    End
End
