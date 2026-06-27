Operation =4
Option =0
Where ="(((Units.UnitTax)<1 Or (Units.UnitTax) Is Null) AND ((Units.UnitStatus)<>\"T\") "
    "AND ((Units.UnitTaxExempt)=False) AND ((Locations.LTaxRate)>0) AND ((CCur(IIf([U"
    "nitTaxExempt]=True,0,IIf([UnitRent]>0,[UnitRent],[UnitSRent])*[LTaxRate])))>0))"
Begin InputTables
    Name ="Units"
    Name ="Locations"
End
Begin OutputColumns
    Name ="Units.UnitTax"
    Expression ="CCur(IIf([UnitTaxExempt]=True,0,IIf([UnitRent]>0,[UnitRent],[UnitSRent])*[LTaxRa"
        "te]))"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Taxes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocTaxRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1249
    Bottom =1021
    Left =-1
    Top =-1
    Right =1173
    Bottom =627
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =250
        Bottom =562
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =399
        Top =294
        Right =543
        Bottom =588
        Top =0
        Name ="Locations"
        Name =""
    End
End
