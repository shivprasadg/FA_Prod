Operation =1
Option =0
Where ="(((MstrLease.MLNo)=forms!Billing!SELLEASE) And ((Schedule.SchID)=forms!Billing!S"
    "ELsCH) And ((Schedule.SchFSL)=Yes) And ((SchGrp.BLCD) Is Not Null) And ((Locatio"
    "ns_2.LocationType)=\"billing\") And ((SchGrp.Assignee_Bank) Is Not Null))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="SchGrp"
    Name ="Clients"
    Alias ="Clients_1"
    Name ="Locations"
    Alias ="Locations_2"
    Name ="Locations"
    Name ="Units"
    Name ="lookup_States"
End
Begin OutputColumns
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.AssignmentDate"
    Expression ="Schedule.SchFSL"
    Expression ="SchGrp.SGrpNO"
    Expression ="Locations.LocationName"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitStatus"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType]"
    Expression ="Units.UnitFinalCost"
    Expression ="Units.UnitTaxExempt"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitPDStartDate"
    Expression ="SchGrp.BLCD_Projected"
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
    Expression ="Locations_2.LocationType"
    Alias ="blOCst2"
    Expression ="Locations_1.LAddrLine1"
    Expression ="lookup_States.StTaxType"
    Expression ="Clients.FAVenorNo"
    Expression ="Schedule.SchPDInv"
    Expression ="Schedule.SchPDInvDt"
    Expression ="Schedule.SchPDBillDt"
    Expression ="Schedule.SchBnkClNO"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Expression ="Units.UnitAcceptDate"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="Units.UNITSRENT"
    Expression ="Units.UnitTax"
    Expression ="Clients_1.clientCompanyName"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePymtFreq"
    Expression ="SchGrp.LeasePaidAdvArr"
    Expression ="SchGrp.SgrpBnkNo"
    Expression ="Schedule.SchBnkClNO"
    Expression ="Clients.ClientShNm"
    Alias ="TaxAmt"
    Expression ="IIf([unittaxexempt]=True,0,[locations].[LTaxRate]*[unitsrent])"
    Expression ="Locations.LTaxRate"
    Expression ="Units.UNITSRENT"
    Expression ="Clients_1.clientAddress1"
    Expression ="Clients_1.clientAddress2"
    Expression ="Clients_1.clientCity"
    Expression ="Clients_1.clientState"
    Expression ="Clients_1.clientZip"
    Expression ="Locations_2.LocationType"
    Expression ="Locations_2.LocationStreet"
    Expression ="Locations_2.LAddrLine1"
    Expression ="Locations_2.LAddrLine2"
    Expression ="Locations_2.LCity"
    Expression ="Locations_2.LState"
    Expression ="Locations_2.LZip"
    Alias ="1stBankRnt"
    Expression ="IIf(IsNull([AssignmentDate]),[BLCD],[AssignmentDate])"
    Alias ="No Pmts"
    Expression ="[LeaseTermPrimary]+DateDiff(\"m\",IIf(IsNull([AssignmentDate]),[BLCD],[Assignmen"
        "tDate]),[BLCD])"
    Expression ="Units.unitstatus"
    Alias ="FSLrent"
    Expression ="IIf([unitstatus]=\"o\" Or [unitstatus]=\"s\" Or [unitstatus]=\"r\",0,IIf(IsNull("
        "[unitsrent]),[unitrent],[unitsrent]))"
    Alias ="FSLServ"
    Expression ="IIf([unitstatus]=\"o\" Or [unitstatus]=\"r\" Or [unitstatus]=\"s\",0,[SgrpFSLPmt"
        "])"
    Expression ="Units.unittax"
    Expression ="Units.unitrent"
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
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Clients_1"
    Expression ="SchGrp.Assignee_Bank = Clients_1.ClientShNm"
    Flag =2
    LeftTable ="Clients_1"
    RightTable ="Locations_2"
    Expression ="Clients_1.clientID = Locations_2.ClientREF"
    Flag =2
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Locations.LocationName"
    Flag =0
    Expression ="Units.unitunitnum"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
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
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3735"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Locations.LocationCounty"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Locations.LocTaxRt"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="40"
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNegBLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBLCD"
        dbInteger "ColumnWidth" ="1320"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="BLocSt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="BLoST2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="BLocZip"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="BLocCity"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="BLocStreet"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="blOCst2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="Clients.FAVenorNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="Schedule.SchPDInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="Schedule.SchPDInvDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="Schedule.SchPDBillDt"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="Schedule.SchBnkClNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="39"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="41"
    End
    Begin
        dbText "Name" ="BLoST3"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Units.UNITSRENT"
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
        dbInteger "ColumnOrder" ="44"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
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
        dbText "Name" ="Locations_2.LocSt2"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LocSt3"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1stBankRnt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="No Pmts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="FSLrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FSLServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Expr1028"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1035"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1039"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpba"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Expr1040"
        dbInteger "ColumnOrder" ="42"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1009"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1243
    Bottom =773
    Left =-1
    Top =-1
    Right =1211
    Bottom =291
    Left =212
    Top =0
    ColumnsShown =539
    Begin
        Left =-209
        Top =8
        Right =-29
        Bottom =152
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =-8
        Top =7
        Right =136
        Bottom =151
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =157
        Top =6
        Right =301
        Bottom =160
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =40
        Top =240
        Right =227
        Bottom =403
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =343
        Top =13
        Right =487
        Bottom =144
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =432
        Top =182
        Right =576
        Bottom =339
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =671
        Top =184
        Right =815
        Bottom =328
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =741
        Top =9
        Right =885
        Bottom =153
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =544
        Top =13
        Right =710
        Bottom =150
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =900
        Top =12
        Right =1044
        Bottom =156
        Top =0
        Name ="lookup_States"
        Name =""
    End
End
