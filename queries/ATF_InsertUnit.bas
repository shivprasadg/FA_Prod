Operation =3
Name ="Units"
Option =2
Where ="(((ATFUnitSpec.ATFUnitSpecID)=[UnitSpecX]) AND ((SchGrp.SGrpID)=[GroupIDx]))"
Begin InputTables
    Name ="ATFUnitSpec"
    Name ="SchGrp"
End
Begin OutputColumns
    Name ="ATFGroupSpecID"
    Expression ="ATFUnitSpec.ATFGroupSpecID"
    Name ="ATFUnitSpecID"
    Expression ="ATFUnitSpec.ATFUnitSpecID"
    Name ="Axle"
    Expression ="ATFUnitSpec.Axle"
    Name ="UnitAxleProperty"
    Expression ="ATFUnitSpec.AxleProperty"
    Name ="ClientID"
    Expression ="ATFUnitSpec.ClientID"
    Name ="ClientGroupID"
    Expression ="ATFUnitSpec.ClientGroupID"
    Name ="UnitDeliveryLocationRef"
    Expression ="ATFUnitSpec.DeliveryLocationID"
    Name ="UnitBuildDate"
    Expression ="ATFUnitSpec.EstimatedBuildDate"
    Name ="UnitEstDelDate"
    Expression ="ATFUnitSpec.EstimatedDeliveryDate"
    Name ="FirstExtRent"
    Expression ="ATFUnitSpec.FirstExtendedRent"
    Name ="UnitGVW"
    Expression ="ATFUnitSpec.GVW"
    Name ="UnitGarageLocationRef"
    Expression ="ATFUnitSpec.GarageLocationID"
    Name ="LegalDescription"
    Expression ="ATFUnitSpec.LegalDescription"
    Name ="MTMRent"
    Expression ="ATFUnitSpec.MTMRent"
    Name ="UNITPDRENT"
    Expression ="ATFUnitSpec.PerDiemRent"
    Name ="UnitRegLocationRef"
    Expression ="ATFUnitSpec.RegTitleLocationID"
    Name ="UnitSalesResid"
    Expression ="ATFUnitSpec.SalesResidual"
    Name ="UNITSRENT"
    Expression ="ATFUnitSpec.SchRent"
    Name ="GroupREF"
    Expression ="SchGrp.SGrpID"
    Name ="UnitAddDt"
    Expression ="ATFUnitSpec.SpecInserted"
    Name ="UNITRENT"
    Expression ="ATFUnitSpec.SyndRent"
    Name ="UnitSynResid"
    Expression ="ATFUnitSpec.SyndResidual"
    Name ="UnitTransmissionType"
    Expression ="ATFUnitSpec.TransmissionType"
    Name ="UnitConfiguration"
    Expression ="ATFUnitSpec.UnitConfiguration"
    Name ="UnitFinalCost"
    Expression ="ATFUnitSpec.UnitCost"
    Name ="UnitDescriptor"
    Expression ="ATFUnitSpec.UnitDescriptor"
    Name ="UnitMake"
    Expression ="ATFUnitSpec.UnitMake"
    Name ="UnitModel"
    Expression ="ATFUnitSpec.UnitModel"
    Name ="UnitYr"
    Expression ="ATFUnitSpec.UnitModelYear"
    Alias ="UnitStatus"
    Name ="UnitStatus"
    Expression ="'B'"
    Name ="UnitSubType"
    Expression ="ATFUnitSpec.UnitSubType"
    Name ="UnitType"
    Expression ="ATFUnitSpec.UnitType"
    Name ="UnitUOM"
    Expression ="ATFUnitSpec.UOM"
    Name ="VendorREF"
    Expression ="ATFUnitSpec.VendorID"
End
Begin Parameters
    Name ="GroupIDx"
    Flag =4
    Name ="UnitSpecX"
    Flag =4
End
Begin Joins
    LeftTable ="ATFUnitSpec"
    RightTable ="SchGrp"
    Expression ="ATFUnitSpec.ATFGroupSpecID = SchGrp.ATFGroupSpecID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ATFUnitSpec.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.BillingContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitConfiguration"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.AxleProperty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitDescriptor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.GVW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.LegalDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.FirstExtendedTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.GarageLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.RegTitleLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.DeliveryLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SalesResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.FirstExtendedRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SyndResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.EstimatedBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.EstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.PerDiemRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SpecInserted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitVin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.TransmissionType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =80
    Top =61
    Right =2219
    Bottom =1047
    Left =-1
    Top =-1
    Right =2115
    Bottom =488
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =56
        Top =63
        Right =315
        Bottom =397
        Top =0
        Name ="ATFUnitSpec"
        Name =""
    End
    Begin
        Left =401
        Top =92
        Right =545
        Bottom =532
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
