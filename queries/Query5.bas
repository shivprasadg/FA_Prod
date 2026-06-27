dbMemo "SQL" ="SELECT R.*, U.UnitID, U.UnitUnitNum AS UnitNum, U.UnitVIN, U.UnitStatus, U.UnitO"
    "ffLeaseDt, U.LocationParkedID, U.ClientGroupID, SG.AssigneeBankID, S.SchID, S.Sc"
    "hNo, RMKTUnitTitleLog_Work.LastEntry\015\012FROM ((((RMKT AS R INNER JOIN Buyers"
    "Order_MMod_Work AS BO ON R.UnitRef = BO.UnitID) INNER JOIN Units AS U ON R.UnitR"
    "ef = U.UnitID) INNER JOIN SchGrp AS SG ON U.GroupREF = SG.SGrpID) INNER JOIN Sch"
    "edule AS S ON SG.SGrpSchID = S.SchID) INNER JOIN RMKTUnitTitleLog_Work ON BO.Uni"
    "tID = RMKTUnitTitleLog_Work.AssetID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="R.AdjBookValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.MileageProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.FreeStrexp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.BuyDateEffective"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.SUF Returned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.BankNotified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktAdjNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.BuyOutCostNet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.PONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.BuyOutCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.PayOffRequested"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.BuyerRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.MiParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.BankShareNetProceeds"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.TargetSaleDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.Commision"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.TermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.Custbookgoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.TitleOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktDeposit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.FAPortfolio"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.DrawDownID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktDepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.DiscountPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.FAAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.FABrkComm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.InvoiceRequested"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.FATitleBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.IFPercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.FMVValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.IFPaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.FreeStrDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.GainLoss"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.HrParked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.IFBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.Shipping"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.IFLoanNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.IFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.SUF Sent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.IFAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.InvoiceRequestNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.TargetSale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.ParkDateProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.ParkDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.PBOS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktBOSSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmKTBOSRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktTitleRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktSPRAdj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmkSm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmkLOSupport"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.RmktTitleSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.SoldPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.SellerInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.WhSValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.Wholesale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.PayOffReceived"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="R.SufPDFLink"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.UnitOffLeaseDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.LocationParkedID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SG.AssigneeBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LT.LastEntry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RMKTUnitTitleLog_Work.LastEntry"
        dbInteger "ColumnWidth" ="2925"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
