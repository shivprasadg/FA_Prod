Operation =1
Option =0
Begin InputTables
    Name ="UnitsFormFilter"
    Name ="Units"
    Name ="SELECT LocationId, LTaxRate, \"<B>\" & LocationStreet & \"</b><Br />\" & LAddrLi"
        "ne1 & \"<Br />\" & IIf(IsNull(LocSubLessee),Null,LocSubLessee & \"<Br />\") & [L"
        "City] & \", \" & [LState] & \" \" & [Lzip] & \" |<b> County</b> \" & LCounty AS "
        "GarAddr, LState, LTaxOption FROM Locations"
    Alias ="Gar"
    Name ="SELECT LocationId, LTaxRate, \"<B>\" & LocationStreet & \"</b><Br />\" & LAddrLi"
        "ne1 & \"<Br />\" & IIf(IsNull(LocSubLessee),Null,LocSubLessee & \"<Br />\") & [L"
        "City] & \", \" & [LState] & \" \" & [Lzip] & \" | <b> County</b> \" & LCounty AS"
        " DelvAddr FROM Locations"
    Alias ="Delv"
    Name ="SELECT LocationId, LTaxRate, \"<B>\" & LocationStreet & \"</b><Br />\" & LAddrLi"
        "ne1 & \"<Br />\" & IIf(IsNull(LocSubLessee),Null,LocSubLessee & \"<Br />\") & [L"
        "City] & \", \" & [LState] & \" \" & [Lzip] & \" | <b> County</b> \" & LCounty AS"
        " TitleAddr, LState FROM Locations"
    Alias ="Title"
    Name ="SELECT LocationId, LTaxRate, \"<B>\" & LocationStreet & \"</b><Br />\" & LAddrLi"
        "ne1 & \"<Br />\" & IIf(IsNull(LocSubLessee),Null,LocSubLessee & \"<Br />\") & [L"
        "City] & \", \" & [LState] & \" \" & [Lzip] & \" |<b> County</b> \" & LCounty AS "
        "RegAddr FROM Locations"
    Alias ="Reg"
    Name ="LookUp_States"
    Name ="Rmkt"
End
Begin OutputColumns
    Expression ="Units.*"
    Expression ="Delv.DelvAddr"
    Expression ="Gar.GarAddr"
    Expression ="Title.TitleAddr"
    Expression ="Reg.RegAddr"
    Alias ="TitleState"
    Expression ="Title.LState"
    Expression ="LookUp_States.StTaxType"
    Expression ="Gar.LTaxOption"
    Expression ="Gar.LTaxRate"
    Expression ="UnitsFormFilter.*"
    Expression ="Rmkt.BuyDateEffective"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Gar"
    Expression ="Units.UnitGarageLocationREF = Gar.LocationID"
    Flag =2
    LeftTable ="UnitsFormFilter"
    RightTable ="Units"
    Expression ="UnitsFormFilter.AssetIDx = Units.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Delv"
    Expression ="Units.UnitDeliveryLocationRef = Delv.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Title"
    Expression ="Units.UnitTitleLocationRef = Title.LocationID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Reg"
    Expression ="Units.UnitRegLocationRef = Reg.LocationID"
    Flag =2
    LeftTable ="Gar"
    RightTable ="LookUp_States"
    Expression ="Gar.LState = LookUp_States.StateAbbrev"
    Flag =2
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
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
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.TUnitLessorOwner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.TitleSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AssetID_FATitleLog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DivisionOwnerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.AcqMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MileageAllowance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOdometer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1170"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.AcqHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDep"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1185"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.PLBReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.CACarbCompDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientISD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.CofASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIRPAcct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitIRP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.delstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.FirstExtRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.DrivingMpgGoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ObcUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorREF"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitDel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitSynResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.MpgGoal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PerDiemSurrender"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportStartDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitTitlingTrust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.RVGuar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOriginserv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAxleProperty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVINLocked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAddDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFASerial"
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
        dbText "Name" ="Units.UnitCofACopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepPdDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitConfiguration"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDepdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDocAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDescriptor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.LegalDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitExckPortRpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitGVW"
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
        dbText "Name" ="Units.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInvoiceDate"
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
        dbText "Name" ="Units.UnitEngineType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRAFSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
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
        dbText "Name" ="Units.UnitPlateState"
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
        dbText "Name" ="Units.UnitPayOff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPilot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPOA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegExp"
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
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTASent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleAtLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleComments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleOwner"
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
        dbText "Name" ="Units.UnitTransmissionType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransSerial"
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
        dbText "Name" ="Units.UnitTaxNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitSalesResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVINVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVendorPytDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Delv.DelvAddr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reg.RegAddr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gar.GarAddr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title.TitleAddr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.ProjOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.LeaseExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.sGrpNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.UnitOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.AssetIDx"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.index"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LookUp_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gar.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gar.LTaxOption"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsFormFilter.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.BuyDateEffective"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =209
    Top =90
    Right =1341
    Bottom =1132
    Left =-1
    Top =-1
    Right =1108
    Bottom =457
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =85
        Top =13
        Right =202
        Bottom =255
        Top =0
        Name ="UnitsFormFilter"
        Name =""
    End
    Begin
        Left =227
        Top =12
        Right =371
        Bottom =531
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =451
        Top =130
        Right =595
        Bottom =242
        Top =0
        Name ="Gar"
        Name =""
    End
    Begin
        Left =451
        Top =12
        Right =595
        Bottom =122
        Top =0
        Name ="Delv"
        Name =""
    End
    Begin
        Left =825
        Top =112
        Right =969
        Bottom =253
        Top =0
        Name ="Title"
        Name =""
    End
    Begin
        Left =452
        Top =357
        Right =596
        Bottom =466
        Top =0
        Name ="Reg"
        Name =""
    End
    Begin
        Left =643
        Top =12
        Right =787
        Bottom =241
        Top =0
        Name ="LookUp_States"
        Name =""
    End
    Begin
        Left =624
        Top =336
        Right =768
        Bottom =480
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
