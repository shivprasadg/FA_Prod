Operation =1
Option =0
Where ="(((vw_CofABase.SGrpID)=[Forms]![BulkChange_Form]![GrpID]))"
Begin InputTables
    Name ="vw_CofABase"
End
Begin OutputColumns
    Expression ="vw_CofABase.*"
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
        dbText "Name" ="vw_CofABase.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.CorpState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientCoNm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.CofASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SchedDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLMo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.GrpCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLDay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MakeModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitTransDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.DescriptionLong"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SubLeasee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.PLBMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.PLBHours"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =224
    Top =102
    Right =1426
    Bottom =1079
    Left =-1
    Top =-1
    Right =1184
    Bottom =683
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =68
        Top =43
        Right =453
        Bottom =592
        Top =0
        Name ="vw_CofABase"
        Name =""
    End
End
