Operation =1
Option =0
Where ="(((SchGrp.SGrpID)=forms!groupform!sgrpid) And ((Units.Unitstatus)=\"B\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Mileage"
End
Begin OutputColumns
    Expression ="Units.UnitTitleLocationRef"
    Expression ="Units.UnitGarageLocationRef"
    Expression ="Units.UnitRegLocationRef"
    Expression ="Units.UnitDeliveryLocationRef"
    Expression ="Units.UnitType"
    Expression ="Units.UnitSubType"
    Expression ="Units.CofASigner"
    Expression ="MstrLease.MLID"
    Expression ="Schedule.SchID"
    Expression ="Units.UnitModel"
    Expression ="Schedule.SchedDt"
    Expression ="Units.UnitID"
    Expression ="Units.unitvin"
    Alias ="GrpCD"
    Expression ="[sgrpno] & [unitdeliverylocationref] & [unitgaragelocationref] & [unittitlelocat"
        "ionref] & [unitregLocationref] & [cofasigner]"
    Expression ="Units.UNITMAKE"
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
    Expression ="SchGrp.SGrpNO"
    Alias ="Expr1"
    Expression ="Units.PRINTCOFA"
    Alias ="vin4"
    Expression ="Right([unitvin],6)"
    Expression ="Units.unitunitnum"
    Alias ="Desc"
    Expression ="[UnitYr] & \"   \" & [UnitMake] & \"   \" & [UnitModel] & \"   \" & \" \" & [Uni"
        "tSubtype] & \" \" & [UnitType] & \" \" & Nz([AXLE]=\"TBD\")"
    Expression ="Units.UnitAcceptDate"
    Expression ="Mileage.Hours"
    Alias ="AquiMI"
    Expression ="Mileage.MilesLTD"
    Alias ="Mileage"
    Expression ="Mileage.Miles"
    Expression ="Units.Unitstatus"
    Expression ="SchGrp.SGrpSchID"
    Expression ="SchGrp.SGRPPLB"
    Expression ="Units.unitrent"
    Expression ="SchGrp.BLCD_Projected"
    Expression ="Schedule.SchFSL"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="Units.UNITTAX"
    Expression ="Units.UNITSRENT"
    Alias ="Unitcost"
    Expression ="IIf([unitfinalcost]=0,Null,[unitfinalcost])"
    Expression ="Units.UnitTaxExempt"
    Alias ="PLBMileage"
    Expression ="IIf([SGRPPLB]=True,Nz([AcqMileage],0),0)"
    Alias ="PLB"
    Expression ="SchGrp.SGRPPLB"
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
    RightTable ="Mileage"
    Expression ="Units.UnitID = Mileage.UnitID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.clientCompanyName"
    Flag =0
    Expression ="Right([unitvin],6)"
    Flag =0
    Expression ="Units.unitunitnum"
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
    End
    Begin
        dbText "Name" ="Clients.clientCoNm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientDBA"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="Clients.clientAddress2"
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
        dbText "Name" ="Clients.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLMO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="mlday"
        dbLong "AggregateType" ="-1"
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
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.CofASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.PRINTCOFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrpCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage.Hours"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchedDt"
        dbLong "AggregateType" ="-1"
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
    End
    Begin
        dbText "Name" ="Units.UNITTAX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITSRENT"
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
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AquiMI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLBMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PRINTCOFA"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-68
    Top =110
    Right =1190
    Bottom =876
    Left =-1
    Top =-1
    Right =1234
    Bottom =319
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =14
        Top =7
        Right =147
        Bottom =210
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =169
        Top =132
        Right =313
        Bottom =276
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =320
        Top =26
        Right =464
        Bottom =307
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =489
        Top =22
        Right =643
        Bottom =300
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =660
        Top =16
        Right =819
        Bottom =267
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =855
        Top =137
        Right =1143
        Bottom =339
        Top =0
        Name ="Mileage"
        Name =""
    End
End
