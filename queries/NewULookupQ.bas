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
    Name ="Vendors"
    Name ="Emps"
End
Begin OutputColumns
    Expression ="Units.*"
    Alias ="CompanyNm"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.ClientShNm"
    Expression ="Clients.clientCity"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Clients.clientState"
    Expression ="MstrLease.MLNo"
    Alias ="SalespersonName"
    Expression ="Emps.EmpNm"
    Expression ="Vendors.VendorName"
    Expression ="Vendors.VendorSNm"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Clients.clientID"
    Alias ="DSUB"
    Expression ="Locations.LocSubLessee"
    Alias ="DST1"
    Expression ="Locations.LocationStreet"
    Alias ="Dst2"
    Expression ="Locations.LAddrLine1"
    Alias ="Dcity"
    Expression ="Locations.LCity"
    Alias ="DCounty"
    Expression ="Locations.LCounty"
    Alias ="Dst"
    Expression ="Locations.LState"
    Alias ="Dzip"
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
    Expression ="Clients.clientID"
    Expression ="MstrLease.Client"
    Expression ="MstrLease.MLID"
    Expression ="Schedule.SchMLRef"
    Expression ="Schedule.SchID"
    Expression ="SchGrp.SGrpSchID"
    Expression ="SchGrp.SGrpID"
    Alias ="GroupR"
    Expression ="Units.GroupREF"
    Alias ="UnitPlateNumH"
    Expression ="IIf([UnitPlateNum] Is Null,\"\",Replace(Replace([UnitPlateNum],\"-\",\"\"),\" \""
        ",\"\"))"
    Alias ="SPInitials"
    Expression ="Emps.EmpInt"
    Expression ="Clients.ClientGroupID"
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
    LeftTable ="Units"
    RightTable ="Vendors"
    Expression ="Units.VendorREF = Vendors.VendorID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Clients"
    RightTable ="Emps"
    Expression ="Clients.LOinitial = Emps.EmpInt"
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
        dbText "Name" ="Units.UnitPilot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitVINLocked"
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
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCOASent"
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
        dbText "Name" ="Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Units.UnitTitleComments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitlingTrust"
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
        dbText "Name" ="Units.unitbodt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.bocost"
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
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCurRentExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2835"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DSUB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DST1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Dst2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dcity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dst"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dzip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rs1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4560"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="MstrLease.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchMLRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompanyNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitPlateNumH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorSNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVINVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitCofACopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDeliveryLocationRef"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Units.UnitTransDays"
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
        dbText "Name" ="Units.UnitMSOCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AppSentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOLInd"
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
        dbText "Name" ="Units.UDrawdonw#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITPMTFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalespersonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SPInitials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[CLIENT GROUP]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =83
    Top =343
    Right =1279
    Bottom =937
    Left =-1
    Top =-1
    Right =1164
    Bottom =479
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =225
        Bottom =564
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =255
        Top =49
        Right =399
        Bottom =193
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =609
        Top =9
        Right =780
        Bottom =148
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =443
        Top =108
        Right =587
        Bottom =252
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =685
        Top =196
        Right =900
        Bottom =340
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =441
        Top =261
        Right =593
        Bottom =405
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =723
        Top =196
        Right =867
        Bottom =340
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =827
        Top =235
        Right =971
        Bottom =379
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =985
        Top =78
        Right =1129
        Bottom =222
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =737
        Top =390
        Right =881
        Bottom =534
        Top =0
        Name ="Emps"
        Name =""
    End
End
