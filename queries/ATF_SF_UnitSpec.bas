Operation =1
Option =2
Begin InputTables
    Name ="ATFGroupSpec"
    Name ="ATF"
    Name ="ATFScheduleSpec"
    Name ="vw_ATF_Avail_Signed_FAID"
End
Begin OutputColumns
    Expression ="vw_ATF_Avail_Signed_FAID.clientID"
    Expression ="vw_ATF_Avail_Signed_FAID.ClientGroupId"
    Expression ="ATFScheduleSpec.ATFSchSpecID"
    Expression ="ATFGroupSpec.ATFGroupSpecID"
    Expression ="ATF.FAID"
    Expression ="ATF.ATFID"
    Alias ="ATFUnitSpecDesc"
    Expression ="vw_ATF_Avail_Signed_FAID.[Spec_Name]"
    Alias ="NumUnitsInSpec"
    Expression ="vw_ATF_Avail_Signed_FAID.Spec_UnitQty"
    Alias ="UnitModelYear"
    Expression ="CInt([Units_UnitYr])"
    Alias ="UnitMake"
    Expression ="vw_ATF_Avail_Signed_FAID.Units_UnitMake"
    Alias ="UnitModel"
    Expression ="vw_ATF_Avail_Signed_FAID.Units_UnitModel"
    Alias ="UnitType"
    Expression ="vw_ATF_Avail_Signed_FAID.Units_UnitType"
    Alias ="UnitSubType"
    Expression ="Nz([Units_UnitSubtype],\"TBD\")"
    Alias ="UnitConfiguration"
    Expression ="vw_ATF_Avail_Signed_FAID.Units_UnitConfiguration"
    Alias ="Axle"
    Expression ="vw_ATF_Avail_Signed_FAID.Units_Axle"
    Alias ="PowerType"
    Expression ="'Diesel'"
    Alias ="EngineType"
    Expression ="vw_ATF_Avail_Signed_FAID.Units_UnitEngineType"
    Alias ="TransmissionType"
    Expression ="vw_ATF_Avail_Signed_FAID.Units_UnitTransmissionType"
    Alias ="GVW"
    Expression ="CInt(Nz([Units_Weight],0))"
    Expression ="vw_ATF_Avail_Signed_FAID.VendorID"
    Alias ="UnitCost"
    Expression ="CCur(Nz([Total_Unit_Price__c],0))"
    Alias ="SchRent"
    Expression ="CCur(Nz([Payment__c],0))"
    Alias ="FirstExtendedRent"
    Expression ="CCur(Nz([Payment__c],0)*0.8)"
    Alias ="SalesResidual"
    Expression ="CCur(Nz([Units_UnitSalesResid],0))"
    Alias ="UTARebate"
    Expression ="CCur(Nz([UTA_Amount_per_Unit__c],0))"
    Alias ="EstimatedDeliveryDate"
    Expression ="CDate(Nz([Units_UnitEstDelDate],\"1/1/1900\"))"
    Alias ="NumUTAs"
    Expression ="Nz([Total_UTAs_per_ERA__c],0)"
    Alias ="LOFee"
    Expression ="CCur(Nz([Fee_Dollar__c],0))"
    Alias ="LOFeePct"
    Expression ="Nz(vw_ATF_Avail_Signed_FAID.Fee__c,0)"
    Alias ="LOFeeGroup"
    Expression ="CCur(Nz([Fee_Dollar__c],0)*[Spec_UnitQty])"
    Alias ="DeliveryLocationID"
    Expression ="0"
    Alias ="GarageLocationID"
    Expression ="0"
    Alias ="RegTitleLocationID"
    Expression ="0"
End
Begin Joins
    LeftTable ="ATF"
    RightTable ="ATFScheduleSpec"
    Expression ="ATF.ATFID = ATFScheduleSpec.ATFID"
    Flag =1
    LeftTable ="ATFGroupSpec"
    RightTable ="ATFScheduleSpec"
    Expression ="ATFGroupSpec.ATFSchSpecID = ATFScheduleSpec.ATFSchSpecID"
    Flag =1
    LeftTable ="ATFGroupSpec"
    RightTable ="vw_ATF_Avail_Signed_FAID"
    Expression ="ATFGroupSpec.ATFGroupSpecDesc = vw_ATF_Avail_Signed_FAID.GroupName"
    Flag =1
    LeftTable ="ATFScheduleSpec"
    RightTable ="vw_ATF_Avail_Signed_FAID"
    Expression ="ATFScheduleSpec.ATFSchName = vw_ATF_Avail_Signed_FAID.SchName"
    Flag =1
    LeftTable ="ATFGroupSpec"
    RightTable ="vw_ATF_Avail_Signed_FAID"
    Expression ="ATFGroupSpec.GrpFAID = vw_ATF_Avail_Signed_FAID.FAID"
    Flag =1
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
        dbText "Name" ="UnitSubType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.ClientGroupId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="UnitModelYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="UnitMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Axle"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="UnitModel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="UnitConfiguration"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="UTARebate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="EngineType"
        dbInteger "ColumnWidth" ="2475"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="NumUTAs"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="SalesResidual"
        dbInteger "ColumnWidth" ="2385"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="SchRent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="LOFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LOFeePct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LOFeeGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="GarageLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeliveryLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstExtendedRent"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.ATFID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="ATFUnitSpecDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="NumUnitsInSpec"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PowerType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GVW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EstimatedDeliveryDate"
        dbInteger "ColumnWidth" ="2415"
        dbInteger "ColumnOrder" ="20"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegTitleLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF.FAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="TransmissionType"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ATFSchSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.VendorID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =110
    Top =663
    Right =1925
    Bottom =1354
    Left =-1
    Top =-1
    Right =1797
    Bottom =209
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =384
        Top =18
        Right =555
        Bottom =298
        Top =0
        Name ="ATFGroupSpec"
        Name =""
    End
    Begin
        Left =34
        Top =2
        Right =178
        Bottom =146
        Top =0
        Name ="ATF"
        Name =""
    End
    Begin
        Left =211
        Top =18
        Right =355
        Bottom =162
        Top =0
        Name ="ATFScheduleSpec"
        Name =""
    End
    Begin
        Left =598
        Top =21
        Right =814
        Bottom =472
        Top =0
        Name ="vw_ATF_Avail_Signed_FAID"
        Name =""
    End
End
