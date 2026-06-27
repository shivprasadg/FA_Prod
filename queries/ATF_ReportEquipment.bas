Operation =1
Option =0
Begin InputTables
    Name ="ATFUnitSpec"
    Name ="Vendors"
    Name ="ATF_ChildPartCost"
End
Begin OutputColumns
    Expression ="ATFUnitSpec.ATFID"
    Alias ="UnitDesc"
    Expression ="[UnitModelYear] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubTyp"
        "e] & \" \" & [UnitType] & \" \" & [Axle]"
    Alias ="Size"
    Expression ="ATFUnitSpec.UOM"
    Expression ="ATFUnitSpec.UnitCost"
    Alias ="Units"
    Expression ="ATFUnitSpec.NumUnitsInSpec"
    Expression ="ATFUnitSpec.Axle"
    Expression ="ATFUnitSpec.VendorID"
    Expression ="Vendors.VendorSNm"
    Expression ="ATFUnitSpec.EstimatedDeliveryDate"
    Expression ="ATFUnitSpec.EstimatedBuildDate"
    Expression ="ATFUnitSpec.SchRent"
    Expression ="ATFUnitSpec.ATFUnitSpecDesc"
    Expression ="ATFUnitSpec.ATFGroupSpecID"
    Expression ="ATFUnitSpec.ATFUnitSpecID"
    Alias ="TotalPartCost"
    Expression ="Sum(CCur(Nz([ChildPartCost],0)))"
End
Begin Joins
    LeftTable ="ATFUnitSpec"
    RightTable ="Vendors"
    Expression ="ATFUnitSpec.VendorID = Vendors.VendorID"
    Flag =2
    LeftTable ="ATFUnitSpec"
    RightTable ="ATF_ChildPartCost"
    Expression ="ATFUnitSpec.ATFUnitSpecID = ATF_ChildPartCost.ATFUnitSpecID"
    Flag =2
End
Begin Groups
    Expression ="ATFUnitSpec.ATFID"
    GroupLevel =0
    Expression ="[UnitModelYear] & \" \" & [UnitMake] & \" \" & [UnitModel] & \" \" & [UnitSubTyp"
        "e] & \" \" & [UnitType] & \" \" & [Axle]"
    GroupLevel =0
    Expression ="ATFUnitSpec.UOM"
    GroupLevel =0
    Expression ="ATFUnitSpec.UnitCost"
    GroupLevel =0
    Expression ="ATFUnitSpec.NumUnitsInSpec"
    GroupLevel =0
    Expression ="ATFUnitSpec.Axle"
    GroupLevel =0
    Expression ="ATFUnitSpec.VendorID"
    GroupLevel =0
    Expression ="Vendors.VendorSNm"
    GroupLevel =0
    Expression ="ATFUnitSpec.EstimatedDeliveryDate"
    GroupLevel =0
    Expression ="ATFUnitSpec.EstimatedBuildDate"
    GroupLevel =0
    Expression ="ATFUnitSpec.SchRent"
    GroupLevel =0
    Expression ="ATFUnitSpec.ATFUnitSpecDesc"
    GroupLevel =0
    Expression ="ATFUnitSpec.ATFGroupSpecID"
    GroupLevel =0
    Expression ="ATFUnitSpec.ATFUnitSpecID"
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
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.EstimatedBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbInteger "ColumnWidth" ="3420"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorSNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.EstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFUnitSpecDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalPartCost"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-463
    Top =59
    Right =1136
    Bottom =1005
    Left =-1
    Top =-1
    Right =1581
    Bottom =618
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =382
        Top =91
        Right =737
        Bottom =270
        Top =0
        Name ="ATFUnitSpec"
        Name =""
    End
    Begin
        Left =820
        Top =250
        Right =964
        Bottom =394
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =1031
        Top =212
        Right =1175
        Bottom =356
        Top =0
        Name ="ATF_ChildPartCost"
        Name =""
    End
End
