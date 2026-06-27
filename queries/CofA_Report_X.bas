Operation =1
Option =0
Where ="(((Units.Unitstatus)=\"A\" Or (Units.Unitstatus)=\"B\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="lookup_States"
    Name ="UnitLastMileageRead"
    Name ="LookUp_UnitDesc"
End
Begin OutputColumns
    Expression ="Units.UnitTitleLocationRef"
    Expression ="Units.UnitGarageLocationRef"
    Expression ="Units.UnitRegLocationRef"
    Expression ="Units.UnitDeliveryLocationRef"
    Alias ="CofASigner"
    Expression ="CDbl(Nz([units].[CofASigner],0))"
    Expression ="MstrLease.MLID"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchedDt"
    Expression ="Units.UnitID"
    Expression ="Units.unitvin"
    Alias ="GrpCD"
    Expression ="[sgrpno] & [unitdeliverylocationref] & [unitgaragelocationref] & [unittitlelocat"
        "ionref] & [unitregLocationref] & [cofasigner]"
    Expression ="Clients.clientID"
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientCoNm2"
    Expression ="Clients.clientDBA"
    Expression ="SchGrp.SGrpID"
    Expression ="Clients.clientAddress1"
    Expression ="Clients.clientAddress2"
    Expression ="Clients.clientCity"
    Expression ="Clients.clientState"
    Expression ="Clients.clientZip"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLDt"
    Alias ="MLYr"
    Expression ="DatePart(\"yyyy\",[mldt])"
    Alias ="MLMo"
    Expression ="MonthName(DatePart(\"m\",[mldt]))"
    Alias ="mlday"
    Expression ="DatePart(\"d\",[mldt])"
    Expression ="Schedule.SchNo"
    Alias ="CorpState"
    Expression ="lookup_States.StateLong"
    Expression ="SchGrp.SGrpNO"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Alias ="ModelYear"
    Expression ="UCase([unityr])"
    Expression ="Units.UnitMake"
    Alias ="MakeModel"
    Expression ="[unitmake] & \"   \" & [unitmodel]"
    Expression ="Units.UnitType"
    Alias ="Type"
    Expression ="Units.UnitType"
    Expression ="Units.UnitSubType"
    Expression ="Units.Axle"
    Alias ="Desc"
    Expression ="IIf([AXLE]=\"TBD\",Null,[AXLE] & \" \") & [UnitSubtype] & \" \" & [UnitType]"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitAcceptDate"
    Expression ="Units.Unitstatus"
    Expression ="SchGrp.SGrpSchID"
    Expression ="SchGrp.SGRPPLB"
    Expression ="Units.UNITRENT"
    Expression ="SchGrp.BLCD_Projected"
    Expression ="Schedule.SchFSL"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="Units.UnitModel"
    Expression ="Units.UNITTAX"
    Alias ="RENT"
    Expression ="IIf([units].[UNITRENT]>0,[units].[UNITRENT],[units].[unitSRent])"
    Alias ="Unitcost"
    Expression ="IIf([unitfinalcost]=0,Null,[unitfinalcost])"
    Expression ="Units.UnitTaxExempt"
    Expression ="Schedule.SubLeasee"
    Expression ="Units.UnitSubType"
    Expression ="UnitLastMileageRead.Mileage"
    Expression ="UnitLastMileageRead.Hours"
    Expression ="UnitLastMileageRead.ReadDate"
    Alias ="DescriptionLong"
    Expression ="Nz([Units].[LegalDescription],[lookup_UnitDesc].[DescriptionLong])"
    Alias ="PLB"
    Expression ="SchGrp.SGRPPLB"
    Alias ="PLBMileage"
    Expression ="IIf([SGRPPLB]=True,Nz([AcqMileage],0),0)"
    Alias ="PLBHours"
    Expression ="IIf([SGRPPLB]=True,Nz([AcqHours],0),0)"
    Alias ="PLBLabel"
    Expression ="Switch([SGRPPLB]=False,Null,Nz([PLBMileage],0)>0,\"Miles as of \" & Format([Form"
        "s]![BulkChange_Form]![PLBAsOfDate],\"mm/dd/yy\"),Nz([PLBHours],0)>0,\"Hours\",Nz"
        "([PLBMileage],0)=0 And Nz([PLBHours],0)=0,\"Miles / Hours\")"
    Expression ="SchGrp.SGRPPLB"
    Alias ="PLBMilesHours"
    Expression ="Switch([SGRPPLB]=False,Null,Nz([PLBMileage],0)>0,Format([PLBMileage],\"#,##0\"),"
        "Nz([PLBHours],0),Format([PLBHours],\"#,##0\"),Nz([PLBMileage],0)>0 And Nz([PLBHo"
        "urs],0)>0,Format([PLBMileage],\"#,##0\") & \" / \" & Format([PLBHours],\"#,##0\""
        "))"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.LeaseTermPrimary"
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
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Clients"
    RightTable ="lookup_States"
    Expression ="Clients.LeaseSTofCorp = lookup_States.StateAbbrev"
    Flag =2
    LeftTable ="Units"
    RightTable ="UnitLastMileageRead"
    Expression ="Units.UnitID = UnitLastMileageRead.UnitID"
    Flag =2
    LeftTable ="Units"
    RightTable ="LookUp_UnitDesc"
    Expression ="Units.UnitID = LookUp_UnitDesc.AssetID"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.clientCompanyName"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
    Expression ="Units.UnitUnitNum"
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Clients.clientCoNm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Clients.clientDBA"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="Clients.clientAddress2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Clients.clientCity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="Clients.clientState"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Clients.clientZip"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="MstrLease.MLDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="MLYr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="MLMO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="mlday"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3405"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="GrpCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Units.Unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTransDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Units.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Schedule.SchedDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITMAKE"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="Units.UNITTAX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unitcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="CorpState"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="MakeModel"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="Units.Axle"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="38"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SubLeasee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofASigner"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLastMileageRead.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLastMileageRead.Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitLastMileageRead.ReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLBMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLBHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLbLabel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="PLBMilesHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.blcd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DescriptionLong"
        dbInteger "ColumnWidth" ="7515"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1035"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1042"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-105
    Top =30
    Right =1126
    Bottom =853
    Left =-1
    Top =-1
    Right =1207
    Bottom =339
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =63
        Top =9
        Right =202
        Bottom =223
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =223
        Top =24
        Right =340
        Bottom =229
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =353
        Top =15
        Right =497
        Bottom =296
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =510
        Top =17
        Right =654
        Bottom =161
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =671
        Top =21
        Right =830
        Bottom =272
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =216
        Top =249
        Right =360
        Bottom =393
        Top =0
        Name ="lookup_States"
        Name =""
    End
    Begin
        Left =867
        Top =182
        Right =1011
        Bottom =326
        Top =0
        Name ="UnitLastMileageRead"
        Name =""
    End
    Begin
        Left =878
        Top =21
        Right =1022
        Bottom =165
        Top =0
        Name ="LookUp_UnitDesc"
        Name =""
    End
End
