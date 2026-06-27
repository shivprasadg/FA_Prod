Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Locations"
    Alias ="Locations_2"
    Name ="lookup_States"
    Name ="Locations"
    Alias ="Locations_3"
    Name ="UnitsFormFilter"
End
Begin OutputColumns
    Expression ="Units.*"
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLOrig"
    Expression ="MstrLease.MLLessor"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Clients.clientID"
    Expression ="Schedule.SchID"
    Expression ="SchGrp.SGrpID"
    Alias ="GSUB"
    Expression ="Locations.LocSubLessee"
    Alias ="GST1"
    Expression ="Locations.LocationStreet"
    Alias ="Gst2"
    Expression ="Locations.LAddrLine1"
    Alias ="Gcity"
    Expression ="Locations.LCity"
    Alias ="GCounty"
    Expression ="Locations.LCounty"
    Alias ="Gst"
    Expression ="Locations.LState"
    Alias ="Gzip"
    Expression ="Locations.LZip"
    Alias ="RSub"
    Expression ="Locations_1.LocSubLessee"
    Alias ="Rs1"
    Expression ="Locations_1.LocationStreet"
    Alias ="RS2"
    Expression ="Locations_1.LAddrLine1"
    Alias ="RCity"
    Expression ="Locations_1.LCity"
    Alias ="RSt"
    Expression ="Locations_1.LState"
    Alias ="Rzip"
    Expression ="Locations_1.LZip"
    Alias ="Tst"
    Expression ="Locations_2.LState"
    Expression ="Locations.LTaxRate"
    Expression ="lookup_States.StTaxType"
    Expression ="SchGrp.BLCD"
    Alias ="DsT1"
    Expression ="Locations_3.LocationStreet"
    Alias ="DsT2"
    Expression ="Locations_3.LAddrLine1"
    Alias ="DsT3"
    Expression ="Locations_3.LAddrLine2"
    Alias ="DCity"
    Expression ="Locations_3.LCity"
    Alias ="DSt"
    Expression ="Locations_3.LState"
    Alias ="DZip"
    Expression ="Locations_3.LZip"
    Alias ="DSub"
    Expression ="Locations_3.LocSubLessee"
    Alias ="GTaxOpt"
    Expression ="Locations.LTaxOption"
    Expression ="Schedule.SchID"
    Alias ="ProjOffLease"
    Expression ="DLookUp(\"OffLeaseProj\",\"vw_sixkeys\",\"AssetID=\" & [UnitID])"
    Alias ="LeaseExp"
    Expression ="DateAdd(\"m\",[SchGrp].[LeaseTermPrimary],[SchGrp].[BLCD])-1"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_1"
    Expression ="Units.UnitRegLocationRef = Locations_1.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_2"
    Expression ="Units.UnitTitleLocationRef = Locations_2.LocationID"
    Flag =2
    LeftTable ="Locations"
    RightTable ="lookup_States"
    Expression ="Locations.LState = lookup_States.StateAbbrev"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations_3"
    Expression ="Units.UnitDeliveryLocationRef = Locations_3.LocationID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="UnitsFormFilter"
    Expression ="Units.UnitID = UnitsFormFilter.AssetIDx"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="GSUB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GST1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gst2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gcity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gst"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gzip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rs1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RS2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rzip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tst"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DsT1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DsT2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DsT3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GTaxOpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
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
        dbText "Name" ="Units.UnitPilot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.EstDelInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
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
        dbText "Name" ="Units.UnitMSOCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AppSentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
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
        dbText "Name" ="Units.UnitMSORecd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleComments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleOwner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPOA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVINLocked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRARSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVINVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PRINTPOFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitlingTrust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFASerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOdometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCurRentExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleAtLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitGVW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITSRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOASent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.CofASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCofACopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
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
        dbText "Name" ="Units.UnitTitleComment"
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
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitLienholder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIRPAcct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTermDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.POASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSalesResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVendorPytDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSynResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIRP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRAFSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPayOff"
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
        dbText "Name" ="Units.UnitOLInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AcqMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.SUFSetn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.SUFRET"
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
        dbText "Name" ="Units.MaintSchA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintUNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLLessor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAddDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.FhutPmtdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOriginserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.FhutReceived"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.CACarbCompDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PakrDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RentTypeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintOther"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.TitleSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.delstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitExckPortRpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPLBPr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AcqHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintProv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.FirstExtRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MaintStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDoubles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepPdDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UDrawdonw#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ObcUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MpgGoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrivingMpgGoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientIDOBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPowerType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDriverType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
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
        dbText "Name" ="Units.ClientISD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DivisionOwnerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DocsSUFSubtitle"
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
        dbText "Name" ="Units.TUnitLessorOwner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAxleProperty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitClientFleetInfo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitConfiguration"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDescriptor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LegalDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstRmktLossAtLPAF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPlateState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTrailerPull"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransmissionType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =133
    Top =-5
    Right =1405
    Bottom =984
    Left =-1
    Top =-1
    Right =1248
    Bottom =604
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =624
        Top =156
        Right =768
        Bottom =300
        Top =0
        Name ="Locations_3"
        Name =""
    End
    Begin
        Left =841
        Top =285
        Right =985
        Bottom =429
        Top =0
        Name ="UnitsFormFilter"
        Name =""
    End
End
