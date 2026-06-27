Operation =1
Option =0
Having ="(((Locations_2.LocationType)=\"billing\") And ((Bills.BillID)=Forms!Billing!Bill"
    "ID) And ((IIf(Units.UnitOffLeasedt<[BillFrom],0,[ActiveRent]))>0) And ((Location"
    "s_2.LocObsolete)=No) And ((SchGrp.BLCD) Is Not Null))"
Begin InputTables
    Name ="Schedule"
    Name ="PDstdtq"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Bills"
    Name ="BillAdjT"
    Name ="SchGrp"
    Name ="Clients"
    Alias ="Clients_1"
    Name ="Locations"
    Alias ="Locations_2"
    Name ="Locations"
    Name ="Units"
    Name ="lookup_States"
    Name ="ExitOptionsActiveRent"
End
Begin OutputColumns
    Expression ="Clients_1.clientID"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Locations.LocationName"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitStatus"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & II"
        "f([UnitType]=\"Tractor\",'',\" \" & [UnitType])"
    Expression ="Units.UnitFinalCost"
    Expression ="Units.UnitTaxExempt"
    Expression ="Locations.LocationStreet"
    Expression ="Locations_2.LocationType"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Units.UnitPDStartDate"
    Expression ="SchGrp.BLCD_Projected"
    Alias ="BLoST3"
    Expression ="Locations_1.LAddrLine2"
    Alias ="BLocStreet"
    Expression ="Locations_1.LocationStreet"
    Alias ="BillToAddress"
    Expression ="[Locations_1].[LocationStreet] & Chr(13) & Chr(10) & [Locations_1].[LAddrLine1] "
        "& Chr(13) & Chr(10) & IIf(IsNull([Locations_1].[LAddrLine2]),Null,[Locations_1]."
        "[LAddrLine2] & Chr(13) & Chr(10)) & [Locations_1].[LCity] & \", \" & [Locations_"
        "1].[LState] & \" \" & Left([Locations_1].[LZip],5) & \"-\" & Right([Locations_1]"
        ".[LZip],4)"
    Alias ="BLoST2"
    Expression ="Locations_1.LAddrLine1"
    Alias ="BLocCity"
    Expression ="Locations_1.LCity"
    Alias ="BLocSt"
    Expression ="Locations_1.LState"
    Alias ="BLocZip"
    Expression ="Locations_1.LZip"
    Alias ="blOCst2"
    Expression ="Locations_1.LAddrLine1"
    Expression ="lookup_States.StTaxType"
    Expression ="Schedule.SchBnkClNO"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Expression ="Units.UnitAcceptDate"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="Units.UNITRENT"
    Expression ="Units.UnitTax"
    Expression ="Clients_1.clientCompanyName"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePymtFreq"
    Expression ="SchGrp.LeasePaidAdvArr"
    Expression ="SchGrp.SgrpBnkNo"
    Expression ="Schedule.SchBnkClNO"
    Expression ="Locations.LTaxRate"
    Expression ="Units.unitoffleasedt"
    Alias ="TaxAmt"
    Expression ="IIf([unittaxexempt]=-1,0,[locations].[LTaxRate]*IIf([units].[unitoffleasedt]<[bi"
        "llfrom],0,[ActiveRent]))"
    Alias ="LineTot"
    Expression ="IIf([Units].[unitoffleasedt]<[billfrom],0,[ActiveRent]+IIf([unittaxexempt]=-1,0,"
        "[locations].[LTaxRate]*[ActiveRent])+Nz([SGrpFSLPmt],0))"
    Expression ="Clients_1.clientAddress1"
    Expression ="Clients_1.clientAddress2"
    Expression ="Clients_1.clientCity"
    Expression ="Clients_1.clientState"
    Expression ="Clients_1.clientZip"
    Expression ="Locations_2.LocationStreet"
    Expression ="Locations_2.LAddrLine1"
    Expression ="Locations_2.LAddrLine2"
    Expression ="Locations_2.LCity"
    Expression ="Locations_2.LState"
    Expression ="Locations_2.LZip"
    Expression ="Bills.BillNo"
    Expression ="Bills.Billdt"
    Expression ="Bills.Billfrom"
    Expression ="Bills.Billto"
    Expression ="Bills.BillDueDt"
    Expression ="Bills.BillID"
    Expression ="PDstdtq.MLNo"
    Alias ="SumOfBillAdj"
    Expression ="Sum(BillAdjT.BillAdj)"
    Expression ="Locations_2.LAddrLine3"
    Alias ="UnitsBillRent"
    Expression ="IIf([Units].[UnitOffLeasedt]<[BillFrom],0,[ActiveRent])"
    Expression ="Units.UCCC"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="PDstdtq"
    Expression ="Schedule.SchID = PDstdtq.SchID"
    Flag =2
    LeftTable ="Schedule"
    RightTable ="Locations_1"
    Expression ="Schedule.SchBillAdd = Locations_1.LocationID"
    Flag =2
    LeftTable ="Schedule"
    RightTable ="Bills"
    Expression ="Schedule.SchID = Bills.BillSchRef"
    Flag =1
    LeftTable ="Bills"
    RightTable ="BillAdjT"
    Expression ="Bills.BillID = BillAdjT.BillAdjBillref"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Clients_1"
    Expression ="SchGrp.Assignee_Bank = Clients_1.ClientShNm"
    Flag =1
    LeftTable ="Clients_1"
    RightTable ="Locations_2"
    Expression ="Clients_1.clientID = Locations_2.ClientREF"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Locations"
    RightTable ="Units"
    Expression ="Locations.LocationID = Units.UnitGarageLocationRef"
    Flag =1
    LeftTable ="Locations"
    RightTable ="lookup_States"
    Expression ="Locations.LState = lookup_States.StateAbbrev"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="ExitOptionsActiveRent"
    Expression ="Units.UnitID = ExitOptionsActiveRent.UnitID"
    Flag =2
End
Begin OrderBy
    Expression ="Locations.LocationName"
    Flag =0
    Expression ="Units.unitunitnum"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
End
Begin Groups
    Expression ="Clients_1.clientID"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Units.unitunitnum"
    GroupLevel =0
    Expression ="Units.UnitVIN"
    GroupLevel =0
    Expression ="Units.UnitStatus"
    GroupLevel =0
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & II"
        "f([UnitType]=\"Tractor\",'',\" \" & [UnitType])"
    GroupLevel =0
    Expression ="Units.UnitFinalCost"
    GroupLevel =0
    Expression ="Units.UnitTaxExempt"
    GroupLevel =0
    Expression ="Locations.LocationStreet"
    GroupLevel =0
    Expression ="Locations_2.LocationType"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LCounty"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Locations.LZip"
    GroupLevel =0
    Expression ="Units.UnitPDStartDate"
    GroupLevel =0
    Expression ="SchGrp.BLCD_Projected"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine2"
    GroupLevel =0
    Expression ="Locations_1.LocationStreet"
    GroupLevel =0
    Expression ="[Locations_1].[LocationStreet] & Chr(13) & Chr(10) & [Locations_1].[LAddrLine1] "
        "& Chr(13) & Chr(10) & IIf(IsNull([Locations_1].[LAddrLine2]),Null,[Locations_1]."
        "[LAddrLine2] & Chr(13) & Chr(10)) & [Locations_1].[LCity] & \", \" & [Locations_"
        "1].[LState] & \" \" & Left([Locations_1].[LZip],5) & \"-\" & Right([Locations_1]"
        ".[LZip],4)"
    GroupLevel =0
    Expression ="Locations_1.LCity"
    GroupLevel =0
    Expression ="Locations_1.LState"
    GroupLevel =0
    Expression ="Locations_1.LZip"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine1"
    GroupLevel =0
    Expression ="lookup_States.StTaxType"
    GroupLevel =0
    Expression ="Right([unitvin],6)"
    GroupLevel =0
    Expression ="Units.UnitAcceptDate"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="Units.UNITRENT"
    GroupLevel =0
    Expression ="Units.UnitTax"
    GroupLevel =0
    Expression ="Clients_1.clientCompanyName"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.LeaseTermPrimary"
    GroupLevel =0
    Expression ="SchGrp.LeasePymtFreq"
    GroupLevel =0
    Expression ="SchGrp.LeasePaidAdvArr"
    GroupLevel =0
    Expression ="SchGrp.SgrpBnkNo"
    GroupLevel =0
    Expression ="Schedule.SchBnkClNO"
    GroupLevel =0
    Expression ="Locations.LTaxRate"
    GroupLevel =0
    Expression ="Units.unitoffleasedt"
    GroupLevel =0
    Expression ="IIf([unittaxexempt]=-1,0,[locations].[LTaxRate]*IIf([units].[unitoffleasedt]<[bi"
        "llfrom],0,[ActiveRent]))"
    GroupLevel =0
    Expression ="IIf([Units].[unitoffleasedt]<[billfrom],0,[ActiveRent]+IIf([unittaxexempt]=-1,0,"
        "[locations].[LTaxRate]*[ActiveRent])+Nz([SGrpFSLPmt],0))"
    GroupLevel =0
    Expression ="Clients_1.clientAddress1"
    GroupLevel =0
    Expression ="Clients_1.clientAddress2"
    GroupLevel =0
    Expression ="Clients_1.clientCity"
    GroupLevel =0
    Expression ="Clients_1.clientState"
    GroupLevel =0
    Expression ="Clients_1.clientZip"
    GroupLevel =0
    Expression ="Locations_2.LocationStreet"
    GroupLevel =0
    Expression ="Locations_2.LAddrLine1"
    GroupLevel =0
    Expression ="Locations_2.LAddrLine2"
    GroupLevel =0
    Expression ="Locations_2.LCity"
    GroupLevel =0
    Expression ="Locations_2.LState"
    GroupLevel =0
    Expression ="Locations_2.LZip"
    GroupLevel =0
    Expression ="Bills.BillNo"
    GroupLevel =0
    Expression ="Bills.Billdt"
    GroupLevel =0
    Expression ="Bills.Billfrom"
    GroupLevel =0
    Expression ="Bills.Billto"
    GroupLevel =0
    Expression ="Bills.BillDueDt"
    GroupLevel =0
    Expression ="Bills.BillID"
    GroupLevel =0
    Expression ="PDstdtq.MLNo"
    GroupLevel =0
    Expression ="Locations_2.LAddrLine3"
    GroupLevel =0
    Expression ="IIf([Units].[UnitOffLeasedt]<[BillFrom],0,[ActiveRent])"
    GroupLevel =0
    Expression ="Units.UCCC"
    GroupLevel =0
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitType] & \" \""
        " & [UnitSubType]"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine1"
    GroupLevel =0
    Expression ="Schedule.SchBnkClNO"
    GroupLevel =0
    Expression ="Locations_2.LocObsolete"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
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
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3735"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLoST2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocCity"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocStreet"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="blOCst2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchBnkClNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLoST3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientCompanyName"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxAmt"
        dbText "Format" ="$#,##0.00;($#,##0.00)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LocationStreet"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billfrom"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDstdtq.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineTot"
        dbText "Format" ="$#,##0.00;($#,##0.00)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfBillAdj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsbillRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1028"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LAddrLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LAddrLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LAddrLine3"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-374
    Top =71
    Right =1108
    Bottom =1022
    Left =-1
    Top =-1
    Right =1458
    Bottom =224
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =339
        Top =7
        Right =483
        Bottom =161
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =513
        Top =216
        Right =694
        Bottom =387
        Top =0
        Name ="PDstdtq"
        Name =""
    End
    Begin
        Left =266
        Top =228
        Right =453
        Bottom =391
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =161
        Top =10
        Right =305
        Bottom =218
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =-12
        Top =203
        Right =132
        Bottom =347
        Top =0
        Name ="BillAdjT"
        Name =""
    End
    Begin
        Left =547
        Top =11
        Right =691
        Bottom =185
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =801
        Top =168
        Right =945
        Bottom =352
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =1010
        Top =171
        Right =1154
        Bottom =347
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =928
        Top =-9
        Right =1072
        Bottom =135
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =701
        Top =1
        Right =867
        Bottom =138
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1227
        Top =4
        Right =1371
        Bottom =148
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =1214
        Top =94
        Right =1359
        Bottom =238
        Top =0
        Name ="ExitOptionsActiveRent"
        Name =""
    End
End
