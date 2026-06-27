dbMemo "SQL" ="SELECT Units.UnitID, Units.UnitVIN, Units.UnitStatus, Units.UnitUnitNum, Units.U"
    "nitUnitNum2, Units.UnitYr, Units.UnitMake, Units.UnitModel, Units.UnitType, Unit"
    "s.Axle, Units.UnitCOASent, Units.UnitPOA, Units.VendorREF, Units.UnitFinalCost, "
    "Units.UnitInvoiceNum, Units.UnitInvoiceDate, Units.UnitVendorPytDate, Units.Unit"
    "InServ, Units.UnitInServInd, Units.UNITPDRENT, Units.UNITRENT, Units.FirstExtRen"
    "t, Units.UnitDeliveryLocationRef, Units.UnitGarageLocationRef, Units.UnitTitleLo"
    "cationRef, Units.UnitRegLocationRef, Units.UnitIRP, Units.UnitIRPAcct, Units.Uni"
    "tPlateNum, Units.UnitRegExp, Units.UnitTax, Units.UnitTaxExempt, Units.UnitTaxNo"
    "tes, Units.UnitRegW, Units.UnitTASent, Units.UnitBuildDate, Units.UnitEstDelDate"
    ", Units.UnitActualDelDate, Units.CofASigner, Units.UnitAcceptDate, Units.UnitDoc"
    "Accpt, Units.UnitDep, Units.UnitDepdt, Units.MaintExp, Units.UnitIFBank, Units.U"
    "nitIFDate, Units.UnitPayOff, Units.UnitIFAmt, Units.DrawDownId, Units.UnitCOABac"
    "k\015\012FROM Units INNER JOIN RMKTBulk_Update_WorkTable ON Units.UnitID=RMKTBul"
    "k_Update_WorkTable.UnitID;\015\012"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Units.UnitCOASent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Units.UnitVendorPytDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Units.UNITPDRENT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="Units.UNITRENT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Units.FirstExtRent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="Units.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="Units.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Units.UnitIRP"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="Units.UnitIRPAcct"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="Units.UnitRegExp"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="Units.UnitTaxNotes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="Units.UnitRegW"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="Units.UnitTASent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="Units.[UnitBuildDate]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="Units.UnitPOA"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="Units.CofASigner"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="39"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbInteger "ColumnOrder" ="38"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
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
        dbText "Name" ="Units.MaintExp"
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
        dbText "Name" ="Units.drawdownid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOABack"
        dbLong "AggregateType" ="-1"
    End
End
