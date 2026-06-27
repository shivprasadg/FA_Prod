Operation =1
Option =0
Begin InputTables
    Name ="ChildParts"
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Alias ="CountOfChildPartID"
    Expression ="Count(ChildParts.ChildPartID)"
    Alias ="PartDescription"
    Expression ="IIf(Nz([ChildType],\"\")=\"\",Null,[ChildType] & \" | \") & IIf(Nz([ChildYear],0"
        ")=0,Null,[ChildYear] & \" \") & IIf(Nz([ChildMake],\"\")=\"\",Null,[ChildMake] &"
        " \" \") & IIf(Nz([ChildClass],\"\")=\"\",Null,[ChildClass] & \" \")"
    Expression ="ChildParts.childFinalCost"
    Expression ="ChildParts.PONumber"
    Alias ="ExtendedCost"
    Expression ="[childFinalCost]*Count([ChildPartID])"
End
Begin Joins
    LeftTable ="ChildParts"
    RightTable ="vw_SixKeys"
    Expression ="ChildParts.PONumber = vw_SixKeys.PurchaseOrder"
    Flag =1
    LeftTable ="ChildParts"
    RightTable ="vw_SixKeys"
    Expression ="ChildParts.UnitREF = vw_SixKeys.UnitID"
    Flag =1
End
Begin Groups
    Expression ="IIf(Nz([ChildType],\"\")=\"\",Null,[ChildType] & \" | \") & IIf(Nz([ChildYear],0"
        ")=0,Null,[ChildYear] & \" \") & IIf(Nz([ChildMake],\"\")=\"\",Null,[ChildMake] &"
        " \" \") & IIf(Nz([ChildClass],\"\")=\"\",Null,[ChildClass] & \" \")"
    GroupLevel =0
    Expression ="ChildParts.childFinalCost"
    GroupLevel =0
    Expression ="ChildParts.PONumber"
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
        dbText "Name" ="ChildParts.childFinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.PONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartDescription"
        dbInteger "ColumnWidth" ="7095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfChildPartID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="7095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildPartID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtendedCost"
        dbInteger "ColumnWidth" ="7095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =271
    Top =43
    Right =1396
    Bottom =709
    Left =-1
    Top =-1
    Right =1107
    Bottom =372
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =329
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =486
        Bottom =373
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
