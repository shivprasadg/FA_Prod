Operation =1
Option =0
Where ="(((BillStEndq.BillID)=Forms!Billing!BillingSubForm.Form!BillID) And ((Units.Unit"
    "InServ) Is Not Null And (Units.UnitInServ)<Bills.Billto))"
Having ="(((Bills.BillID)=[Forms]![Billing]![BillingSubForm].[Form]![BillID]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="Bills"
    Name ="BillStEndq"
    Name ="SchGrp"
    Name ="Locations"
    Name ="Units"
    Name ="lookup_States"
    Name ="INTDaysQ"
    Name ="BillUnitq"
End
Begin OutputColumns
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.SGrpNoSlv"
    Expression ="SchGrp.SGRPNoPD"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitID"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitAcceptDate"
    Alias ="Desc"
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & \" \" & [unittype]"
    Expression ="Units.UnitTaxExempt"
    Expression ="Locations.LocationName"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Alias ="EndDate"
    Expression ="BillStEndq.BillEndDt"
    Alias ="BLoST3"
    Expression ="Locations_1.LAddrLine2"
    Alias ="BLocStreet"
    Expression ="Locations_1.LocationStreet"
    Alias ="BLoST2"
    Expression ="Locations_1.LAddrLine1"
    Alias ="BLocCity"
    Expression ="Locations_1.LCity"
    Alias ="BLocSt"
    Expression ="Locations_1.LState"
    Alias ="BLocZip"
    Expression ="Locations_1.LZip"
    Expression ="Clients.clientCompanyName"
    Alias ="blOCst2"
    Expression ="Locations_1.LAddrLine1"
    Expression ="lookup_States.StTaxType"
    Expression ="Bills.BillNo"
    Expression ="Clients.FAVenorNo"
    Expression ="Schedule.SchBnkClNO"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Expression ="Units.unitpdrent"
    Alias ="BillStdt"
    Expression ="BillStEndq.Billstart"
    Expression ="Bills.BillID"
    Expression ="Units.unitinserv"
    Expression ="Schedule.schcrexp"
    Expression ="INTDaysQ.SlvDays"
    Alias ="SlvRent"
    Expression ="IIf([sgrpnoslv]=No,0,[UnitPDRent]*[slvdays])"
    Alias ="DSlvRent"
    Expression ="IIf([sgrpnoslv]=No,0,(([UnitPDRent]-([SchCurrRent]/30))))"
    Alias ="PdRent"
    Expression ="IIf([SgrpNoPD]=True,0,[UnitPDRent])"
    Alias ="Dailytax"
    Expression ="IIf([Schfsl]=0,[unittax]/30,0)"
    Expression ="Locations.LTaxRate"
    Expression ="Schedule.SchFSL"
    Expression ="BillUnitq.UnitPDST"
    Expression ="Units.unitrent"
    Expression ="Units.unittax"
    Expression ="Bills.BillDueDt"
    Expression ="Bills.DueonRec"
    Expression ="BillStEndq.SlvEnd"
    Expression ="Bills.Billdt"
    Expression ="Bills.BillSent"
    Expression ="INTDaysQ.UBSlvS"
    Alias ="Bill1"
    Expression ="UCase([invfld1desc] & \" \" & [invfld1])"
    Alias ="Bill2"
    Expression ="Nz([invfld2desc] & \" \" & [invfld2],\"\")"
    Alias ="PDDAYS"
    Expression ="IIf([unitpdst]>[billenddt],0,DateDiff(\"d\",[unitpdst],IIf(IsNull([UNITOFfleased"
        "t ]),[billenddt],[UNITOFFLEASEDT]))+1)"
    Expression ="Units.unitoffleasedt"
    Expression ="SchGrp.SGrpFSLPmt"
    Alias ="UnitCount"
    Expression ="IIf([unitpdrent]>0,1,0)"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="Locations_1"
    Expression ="Schedule.SchBillAdd = Locations_1.LocationID"
    Flag =2
    LeftTable ="Schedule"
    RightTable ="Bills"
    Expression ="Schedule.SchID = Bills.BillSchRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="BillStEndq"
    Expression ="Schedule.SchID = BillStEndq.SchID"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
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
    Flag =1
    LeftTable ="Units"
    RightTable ="INTDaysQ"
    Expression ="Units.UnitID = INTDaysQ.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="BillUnitq"
    Expression ="Units.UnitID = BillUnitq.UnitID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="Units.unitunitnum"
    Flag =0
    Expression ="Locations.LocationName"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
End
Begin Groups
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.SGrpNoSlv"
    GroupLevel =0
    Expression ="SchGrp.SGRPNoPD"
    GroupLevel =0
    Expression ="Units.unitunitnum"
    GroupLevel =0
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Units.UnitVIN"
    GroupLevel =0
    Expression ="Units.UnitStatus"
    GroupLevel =0
    Expression ="Units.UnitAcceptDate"
    GroupLevel =0
    Expression ="[unityr] & \" \" & [unitmake] & \" \" & \" \" & [unittype]"
    GroupLevel =0
    Expression ="Units.UnitTaxExempt"
    GroupLevel =0
    Expression ="Locations.LocationName"
    GroupLevel =0
    Expression ="Locations.LCity"
    GroupLevel =0
    Expression ="Locations.LCounty"
    GroupLevel =0
    Expression ="Locations.LState"
    GroupLevel =0
    Expression ="Locations.LZip"
    GroupLevel =0
    Expression ="BillStEndq.BillEndDt"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine2"
    GroupLevel =0
    Expression ="Locations_1.LocationStreet"
    GroupLevel =0
    Expression ="Locations_1.LCity"
    GroupLevel =0
    Expression ="Locations_1.LState"
    GroupLevel =0
    Expression ="Locations_1.LZip"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine1"
    GroupLevel =0
    Expression ="lookup_States.StTaxType"
    GroupLevel =0
    Expression ="Bills.BillNo"
    GroupLevel =0
    Expression ="Clients.FAVenorNo"
    GroupLevel =0
    Expression ="Schedule.SchBnkClNO"
    GroupLevel =0
    Expression ="Right([unitvin],6)"
    GroupLevel =0
    Expression ="Units.unitpdrent"
    GroupLevel =0
    Expression ="BillStEndq.Billstart"
    GroupLevel =0
    Expression ="Bills.BillID"
    GroupLevel =0
    Expression ="Units.unitinserv"
    GroupLevel =0
    Expression ="Schedule.schcrexp"
    GroupLevel =0
    Expression ="INTDaysQ.SlvDays"
    GroupLevel =0
    Expression ="IIf([sgrpnoslv]=No,0,[UnitPDRent]*[slvdays])"
    GroupLevel =0
    Expression ="IIf([sgrpnoslv]=No,0,(([UnitPDRent]-([SchCurrRent]/30))))"
    GroupLevel =0
    Expression ="IIf([SgrpNoPD]=True,0,[UnitPDRent])"
    GroupLevel =0
    Expression ="IIf([Schfsl]=0,[unittax]/30,0)"
    GroupLevel =0
    Expression ="Locations.LTaxRate"
    GroupLevel =0
    Expression ="Schedule.SchFSL"
    GroupLevel =0
    Expression ="BillUnitq.UnitPDST"
    GroupLevel =0
    Expression ="Units.unitrent"
    GroupLevel =0
    Expression ="Units.unittax"
    GroupLevel =0
    Expression ="Bills.BillDueDt"
    GroupLevel =0
    Expression ="Bills.DueonRec"
    GroupLevel =0
    Expression ="BillStEndq.SlvEnd"
    GroupLevel =0
    Expression ="Bills.Billdt"
    GroupLevel =0
    Expression ="Bills.BillSent"
    GroupLevel =0
    Expression ="INTDaysQ.UBSlvS"
    GroupLevel =0
    Expression ="UCase([invfld1desc] & \" \" & [invfld1])"
    GroupLevel =0
    Expression ="Nz([invfld2desc] & \" \" & [invfld2],\"\")"
    GroupLevel =0
    Expression ="IIf([unitpdst]>[billenddt],0,DateDiff(\"d\",[unitpdst],IIf(IsNull([UNITOFfleased"
        "t ]),[billenddt],[UNITOFFLEASEDT]))+1)"
    GroupLevel =0
    Expression ="Units.unitoffleasedt"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="IIf([unitpdrent]>0,1,0)"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine1"
    GroupLevel =0
    Expression ="Bills.Billdt"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
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
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="2835"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3105"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2130"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbInteger "ColumnWidth" ="1875"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocSt"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLoST2"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocZip"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocCity"
        dbInteger "ColumnWidth" ="1770"
        dbInteger "ColumnOrder" ="23"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLocStreet"
        dbInteger "ColumnWidth" ="2985"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="2985"
        dbInteger "ColumnOrder" ="26"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="blOCst2"
        dbInteger "ColumnWidth" ="1995"
        dbInteger "ColumnOrder" ="27"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.FAVenorNo"
        dbInteger "ColumnOrder" ="31"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchBnkClNO"
        dbInteger "ColumnOrder" ="32"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbInteger "ColumnOrder" ="33"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitpdrent"
        dbInteger "ColumnOrder" ="34"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbInteger "ColumnOrder" ="37"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLoST3"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbInteger "ColumnOrder" ="36"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillNo"
        dbInteger "ColumnOrder" ="29"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.Billdt"
        dbInteger "ColumnOrder" ="30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillID"
        dbInteger "ColumnOrder" ="38"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoSlv"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPNoPD"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitinserv"
        dbInteger "ColumnOrder" ="39"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.schcrexp"
        dbInteger "ColumnOrder" ="40"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SlvRent"
        dbInteger "ColumnOrder" ="42"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PdRent"
        dbInteger "ColumnOrder" ="43"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dailytax"
        dbInteger "ColumnOrder" ="44"
        dbText "Format" ="$#,##0.00;($#,##0.00)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbInteger "ColumnOrder" ="46"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStdt"
        dbInteger "ColumnOrder" ="35"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillDueDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EndDate"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INTDaysQ.SlvDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillStEndq.SlvEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DSlvRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillUnitq.UnitPDST"
        dbInteger "ColumnOrder" ="47"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDDAYS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.BillSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="INTDaysQ.UBSlvS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bill1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bill2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITOFFLEASEDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitid"
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
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =80
    Top =105
    Right =1338
    Bottom =751
    Left =-1
    Top =-1
    Right =1234
    Bottom =286
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =19
        Top =4
        Right =199
        Bottom =148
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =213
        Top =1
        Right =357
        Bottom =145
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =364
        Top =-12
        Right =508
        Bottom =142
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =242
        Top =159
        Right =403
        Bottom =369
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =462
        Top =149
        Right =657
        Bottom =367
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =31
        Top =170
        Right =187
        Bottom =387
        Top =0
        Name ="BillStEndq"
        Name =""
    End
    Begin
        Left =528
        Top =-8
        Right =669
        Bottom =136
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =874
        Top =11
        Right =998
        Bottom =92
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =680
        Top =-4
        Right =846
        Bottom =166
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =1021
        Top =9
        Right =1165
        Bottom =88
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =693
        Top =186
        Right =837
        Bottom =394
        Top =0
        Name ="INTDaysQ"
        Name =""
    End
    Begin
        Left =914
        Top =179
        Right =1054
        Bottom =399
        Top =0
        Name ="BillUnitq"
        Name =""
    End
End
