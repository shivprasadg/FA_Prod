Operation =1
Option =0
Where ="(((MstrLease.MLNo)=forms!Billing!SelLease) And ((Schedule.SchID)=forms!Billing!s"
    "elsch) And ((SchGrp.BLCD) Is Not Null) And ((Schedule.SchFSL)=No))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="PDstdtq"
    Name ="Locations"
    Alias ="Locations_1"
    Name ="SchGrp"
    Name ="Clients"
    Alias ="Clients_1"
    Name ="BankBillAddq"
    Name ="Locations"
    Name ="Units"
    Name ="lookup_States"
End
Begin OutputColumns
    Expression ="Schedule.SchID"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
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
    Alias ="St"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.UnitPDStartDate"
    Expression ="SchGrp.BLCD"
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
    Alias ="blOCst2"
    Expression ="Locations_1.LAddrLine1"
    Expression ="lookup_States.StTaxType"
    Expression ="Clients.FAVenorNo"
    Expression ="Schedule.SchBnkClNO"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Expression ="Units.UnitAcceptDate"
    Expression ="SchGrp.SGrpFSLPmt"
    Alias ="URenT"
    Expression ="IIf([unitstatus]=\"O\" Or [unitstatus]=\"r\" Or [unitstatus]=\"s\",0,[unitsrent]"
        ")"
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
    Expression ="IIf([unitstatus]=\"o\" Or [unitstatus]=\"r\" Or [unitstatus]=\"s\",0,IIf([unitta"
        "xexempt]=True,0,[locations].[LTaxRate]*[unitsrent]))"
    Expression ="Locations.LTaxRate"
    Alias ="LocSt"
    Expression ="Locations.LState"
    Expression ="Units.UNITSRENT"
    Expression ="Clients_1.clientAddress1"
    Expression ="Clients_1.clientAddress2"
    Expression ="Clients_1.clientCity"
    Expression ="Clients_1.clientState"
    Expression ="Clients_1.clientZip"
    Expression ="Units.unitstatus"
    Alias ="1BankRent"
    Expression ="IIf(IsNull([AssignmentDate]),[BLCD],[AssignmentDate])"
    Alias ="bloc3"
    Expression ="Locations_1.LAddrLine2"
    Expression ="BankBillAddq.LocationStreet"
    Expression ="BankBillAddq.LAddrLine1"
    Expression ="BankBillAddq.LAddrLine2"
    Expression ="BankBillAddq.LAddrLine3"
    Expression ="BankBillAddq.LCity"
    Expression ="BankBillAddq.LState"
    Expression ="BankBillAddq.LZip"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="PDstdtq"
    Expression ="Schedule.SchID = PDstdtq.SchID"
    Flag =2
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
    RightTable ="BankBillAddq"
    Expression ="Clients_1.clientID = BankBillAddq.clientID"
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
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
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
        dbText "Name" ="Desc"
        dbInteger "ColumnWidth" ="3735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="Locations.LocationCity"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCounty"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
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
    End
    Begin
        dbText "Name" ="Units.UnitPDStartDate"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNegBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBLCD"
        dbInteger "ColumnWidth" ="1320"
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
        dbText "Name" ="Clients.clientCompanyName"
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
        dbText "Name" ="Clients.FAVenorNo"
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
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="URenT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1BankRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1019"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1031"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocSt"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="St"
        dbInteger "ColumnWidth" ="3255"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="bloc3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankBillAddq.LocationStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankBillAddq.LocSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankBillAddq.LocSt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankBillAddq.LocSt4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankBillAddq.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankBillAddq.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BankBillAddq.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UNITSRENT"
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
    Bottom =249
    Left =203
    Top =0
    ColumnsShown =539
    Begin
        Left =-200
        Top =8
        Right =-20
        Bottom =152
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =1
        Top =7
        Right =145
        Bottom =151
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =166
        Top =6
        Right =310
        Bottom =160
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =279
        Top =221
        Right =460
        Bottom =392
        Top =0
        Name ="PDstdtq"
        Name =""
    End
    Begin
        Left =11
        Top =183
        Right =198
        Bottom =349
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =352
        Top =13
        Right =496
        Bottom =144
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =500
        Top =181
        Right =644
        Bottom =365
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =713
        Top =214
        Right =857
        Bottom =358
        Top =0
        Name ="BankBillAddq"
        Name =""
    End
    Begin
        Left =750
        Top =9
        Right =894
        Bottom =153
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =553
        Top =13
        Right =719
        Bottom =150
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =982
        Top =8
        Right =1126
        Bottom =152
        Top =0
        Name ="lookup_States"
        Name =""
    End
End
