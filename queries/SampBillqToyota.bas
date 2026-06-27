Operation =1
Option =0
Having ="(((Units.UnitStatus)=\"a\") And ((Bills.BillID)=Forms!Billing!BillingSubForm.For"
    "m!BillID) And ((Locations_2.LocObsolete)=No) And ((Locations_2.LocationType)=\"b"
    "illing\") And ((SchGrp.BLCD) Is Not Null))"
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
End
Begin OutputColumns
    Expression ="Clients_1.clientID"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="Locations.LocationName"
    Expression ="Units.unitunitnum"
    Expression ="Units.UnitStatus"
    Alias ="Desc"
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubtype] & \""
        " \" & [UnitType]"
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
    Alias ="UnitsBillRent"
    Expression ="Units.UNITSRENT"
    Expression ="Units.UNITSRENT"
    Expression ="Units.UnitTax"
    Alias ="aSSIGNEEnM"
    Expression ="Clients_1.clientCompanyName"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePymtFreq"
    Expression ="SchGrp.LeasePaidAdvArr"
    Expression ="SchGrp.SgrpBnkNo"
    Expression ="Schedule.SchBnkClNO"
    Expression ="Locations.LTaxRate"
    Alias ="TaxAmt"
    Expression ="IIf([unittaxexempt]=True,0,[locations].[LTaxRate]*[unitsrent])"
    Expression ="Units.UNITSRENT"
    Alias ="LineTot"
    Expression ="IIf([unittaxexempt]=True,0,[locations].[LTaxRate]*[unitsrent])+[unitsrent]"
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
    Expression ="Locations_2.LocationType"
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
    Expression ="Locations_2.LocObsolete"
    Expression ="Locations_2.LAddrLine3"
    Expression ="Units.unitvin"
    Expression ="Units.unitmake"
    Expression ="Units.unittype"
    Expression ="SchGrp.AssignmentDate"
    Expression ="Schedule.schbnkclno"
    Alias ="UnitCount"
    Expression ="IIf([UNITSRENT]>0,1,0)"
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
    Expression ="Units.UnitStatus"
    GroupLevel =0
    Expression ="Units.UnitFinalCost"
    GroupLevel =0
    Expression ="Units.UnitTaxExempt"
    GroupLevel =0
    Expression ="Locations.LocationStreet"
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
    Expression ="IIf([unittaxexempt]=True,0,[locations].[LTaxRate]*[unitsrent])"
    GroupLevel =0
    Expression ="Units.UNITSRENT"
    GroupLevel =0
    Expression ="IIf([unittaxexempt]=True,0,[locations].[LTaxRate]*[unitsrent])+[unitsrent]"
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
    Expression ="Locations_2.LocationType"
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
    Expression ="Locations_2.LocObsolete"
    GroupLevel =0
    Expression ="Locations_2.LAddrLine3"
    GroupLevel =0
    Expression ="Units.unitvin"
    GroupLevel =0
    Expression ="Units.unitmake"
    GroupLevel =0
    Expression ="Units.unittype"
    GroupLevel =0
    Expression ="SchGrp.AssignmentDate"
    GroupLevel =0
    Expression ="Schedule.schbnkclno"
    GroupLevel =0
    Expression ="IIf([UNITSRENT]>0,1,0)"
    GroupLevel =0
    Expression ="[UnitYr] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitType] & \" \""
        " & [UnitSubType]"
    GroupLevel =0
    Expression ="Locations_1.LAddrLine1"
    GroupLevel =0
    Expression ="Units.UNITSRENT"
    GroupLevel =0
    Expression ="Units.UNITSRENT"
    GroupLevel =0
    Expression ="Locations_2.LocationType"
    GroupLevel =0
    Expression ="Schedule.schbnkclno"
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
        dbText "Name" ="Locations.stateRef"
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
        dbText "Name" ="Units.UNITSRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
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
        dbInteger "ColumnWidth" ="3165"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LocationCity"
        dbInteger "ColumnWidth" ="2940"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Locations_2.LocObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_2.LocSt4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unittype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="aSSIGNEEnM"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients_1.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsBillRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1011"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1026"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1031"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1039"
        dbInteger "ColumnWidth" ="2970"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =443
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =276
        Top =1
        Right =420
        Bottom =155
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =506
        Top =227
        Right =687
        Bottom =398
        Top =0
        Name ="PDstdtq"
        Name =""
    End
    Begin
        Left =255
        Top =213
        Right =442
        Bottom =376
        Top =0
        Name ="Locations_1"
        Name =""
    End
    Begin
        Left =101
        Top =-7
        Right =245
        Bottom =201
        Top =0
        Name ="Bills"
        Name =""
    End
    Begin
        Left =45
        Top =211
        Right =189
        Bottom =355
        Top =0
        Name ="BillAdjT"
        Name =""
    End
    Begin
        Left =540
        Top =-7
        Right =684
        Bottom =167
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =708
        Top =214
        Right =852
        Bottom =398
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =908
        Top =211
        Right =1052
        Bottom =387
        Top =0
        Name ="Locations_2"
        Name =""
    End
    Begin
        Left =1024
        Top =3
        Right =1168
        Bottom =147
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =756
        Top =13
        Right =922
        Bottom =224
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
End
