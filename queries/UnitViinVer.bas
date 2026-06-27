dbMemo "SQL" ="SELECT Units.*, RmktVinDecodeT.[Engine Manufacturer], RmktVinDecodeT.[Engine Ser"
    "ies], RmktVinDecodeT.[Engine Series], RmktVinDecodeT.[Engine Configuration], Rmk"
    "tVinDecodeT.[Engine Capacity], RmktVinDecodeT.[Engine Capacity], RmktVinDecodeT."
    "[Fuel Type], RmktVinDecodeT.[Fuel Type]\015\012FROM Units LEFT JOIN RmktVinDecod"
    "eT ON Units.UnitID=RmktVinDecodeT.UnitID;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="RmktVinDecodeT.[Engine Manufacturer]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktVinDecodeT.[Engine Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktVinDecodeT.[Engine Configuration]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktVinDecodeT.[Engine Capacity]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktVinDecodeT.[Fuel Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPilot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.EstDelInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitatTrans"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVINLocked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVINVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOdometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOASent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCofACopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOABack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTASent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.CofASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleComment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMSORecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMSOCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AppSentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitlingTrust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleComments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleRecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleAtLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleOwner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitLienholder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitGVW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPOA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.POASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVendorPytDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIRP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIRPAcct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTermDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDSPR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFASerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRAFSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRARSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITSRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPayOff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSalesResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSynResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitNtType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitNotRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitNotSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOLInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCurRentPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PRINTCOFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PRINTPOFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCurRentExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PNUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UBnkAcct#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitbodt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.bocost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.SUFSetn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.SUFRET"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PrintSUF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.InspDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.InspFail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.InspPass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.InspBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PakrDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PortUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.BuyerREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RmktSalePr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RmktSaleDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.TitleSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LesseeAcct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RevDel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.delstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBOsource"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPLBPr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOriginser"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UOIn-ServInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AcqMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AcqHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintProv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintOther"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintSchA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintUNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAddDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.FhutPmtdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.FhutReceived"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintProvHr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintProvMi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintHrChg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.CACarbCompDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.SchARef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDoubles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFTA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTollPass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepPdDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.SpareRental"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.EPowSpecref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1007"
        dbLong "AggregateType" ="-1"
    End
End
