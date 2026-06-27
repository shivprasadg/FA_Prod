Operation =1
Option =0
Where ="(((ChildParts.childFinalCost)>0) AND ((ChildParts.ChildVendorPytDate) Is Null) A"
    "ND ((ChildParts.ExclFrGrpOEC)=False))"
Begin InputTables
    Name ="Vendors"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="ChildParts.ChildMake"
    Expression ="ChildParts.ChildClass"
    Expression ="ChildParts.ChildType"
    Expression ="ChildParts.ChildInvoiceNum"
    Expression ="ChildParts.childFinalCost"
    Expression ="Vendors.VendorName"
    Expression ="ChildParts.ChildVendorPytDate"
End
Begin Joins
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="ChildParts"
    Expression ="Units.UnitID = ChildParts.UnitREF"
    Flag =1
    LeftTable ="Vendors"
    RightTable ="ChildParts"
    Expression ="Vendors.VendorID = ChildParts.ChildVendorREF"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitID"
    Flag =0
    Expression ="ChildParts.ChildType"
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
        dbText "Name" ="ChildParts.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildInvoiceNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.childFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVendorPytDate"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1290
    Bottom =776
    Left =-1
    Top =-1
    Right =1258
    Bottom =169
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1159
        Top =-9
        Right =1369
        Bottom =215
        Top =0
        Name ="Vendors"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =1073
        Bottom =156
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
