Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Bulk_Update_WorkTable"
    Name ="Locations"
End
Begin OutputColumns
    Expression ="Units.*"
    Expression ="Locations.LState"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Bulk_Update_WorkTable"
    Expression ="Units.UnitID = Bulk_Update_WorkTable.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
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
        dbText "Name" ="Units.[unitsubtype]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitsrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMSOCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitConfiguration"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCofACopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDescriptor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAxleProperty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransmissionType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPilot"
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
        dbText "Name" ="Units.UnitCOABack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMSORecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransSerial"
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
        dbText "Name" ="Units.UnitGVW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSynResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
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
        dbText "Name" ="Units.UnitSalesResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AcqHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAddDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOLInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AcqMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.TitleSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.delstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.CACarbCompDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MpgGoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepPdDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitExckPortRpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOriginserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MileageAllowance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ObcUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrivingMpgGoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LegalDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DivisionOwnerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.TUnitLessorOwner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AssetID_FATitleLog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPlateState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PowerType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LocationParkedID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientISD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LienRelease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LocationAliasID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PerDiemSurrender"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PLBReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RVGuar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RVGuarOEM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatusID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.isUpFront"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =68
    Top =20
    Right =1390
    Bottom =948
    Left =-1
    Top =-1
    Right =1298
    Bottom =440
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =229
        Top =18
        Right =593
        Bottom =296
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Bulk_Update_WorkTable"
        Name =""
    End
    Begin
        Left =714
        Top =72
        Right =858
        Bottom =419
        Top =0
        Name ="Locations"
        Name =""
    End
End
