Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="ChildPart_Cost_OEC_summed"
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="UnitGroupID"
    Expression ="Units.GroupREF"
    Alias ="UnitCost"
    Expression ="Sum(Units.UnitFinalCost)"
    Alias ="ChildCost"
    Expression ="Sum(CCur(Nz([ChildPartCost],0)))"
    Alias ="GroupOEC"
    Expression ="Sum(CCur([UnitFinalCost]+Nz([ChildPartCost],0)))"
    Alias ="GroupRent"
    Expression ="CCur(Sum(Nz([UnitRent],0)))"
    Alias ="GroupLRF"
    Expression ="IIf([GroupRent]=0 Or [GroupOEC]=0,0,CDbl([GroupRent]/[GroupOEC]))"
    Alias ="GroupSalesRent"
    Expression ="CCur(Sum(Nz([UnitSRent],0)))"
    Alias ="GroupLoLRF"
    Expression ="IIf([GroupSalesRent]=0 Or [GroupOEC]=0,0,CDbl([GroupSalesRent]/[GroupOEC]))"
    Alias ="GroupExtRent"
    Expression ="IIf(Sum(Nz([FirstExtRent],0))=0,0,CCur(Sum([FirstExtRent])/Nz([UnitsInGroup],1))"
        ")"
    Alias ="GroupMTMRent"
    Expression ="IIf(Sum(Nz([Units.MTMRent]))=0,0,CCur(Sum([Units.MTMRent])/Nz([UnitsInGroup],1))"
        ")"
    Expression ="SchGrp.UnitsInGroup"
    Alias ="MTMLRF"
    Expression ="IIf([GroupMTMRent]=0 Or [GroupOEC]=0,0,CDbl([GroupMTMRent]/([GroupOEC]/[Unitsing"
        "roup])))"
    Alias ="EXTLRF"
    Expression ="IIf([GroupExtRent]=0 Or [GroupOEC]=0,0,CDbl([GroupExtRent]/([GroupOEC]/[UnitsInG"
        "roup])))"
    Alias ="GroupSalesRV"
    Expression ="Sum(Units.UnitSalesResid)"
    Alias ="GroupSyndRV"
    Expression ="Sum(Units.UnitSynResid)"
    Alias ="VendorPlus"
    Expression ="Sum([VendorCost]+Nz([PartsCost],0))"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="ChildPart_Cost_OEC_summed"
    Expression ="Units.UnitID = ChildPart_Cost_OEC_summed.UnitId"
    Flag =2
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =1
End
Begin Groups
    Expression ="Units.GroupREF"
    GroupLevel =0
    Expression ="SchGrp.UnitsInGroup"
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
        dbText "Name" ="ChildCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="GroupRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupOEC"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="GroupLRF"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="Percent"
        dbByte "DecimalPlaces" ="6"
    End
    Begin
        dbText "Name" ="GroupSalesRent"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupLoLRF"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Percent"
        dbByte "DecimalPlaces" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupExtRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupMTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.unitsingroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EXTLRF"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Percent"
        dbByte "DecimalPlaces" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MTMLRF"
        dbText "Format" ="Percent"
        dbByte "DecimalPlaces" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupSalesRV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupSyndRV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VendorPlus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-61
    Top =126
    Right =1517
    Bottom =1072
    Left =-1
    Top =-1
    Right =1554
    Bottom =306
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =342
        Bottom =409
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =571
        Top =90
        Right =820
        Bottom =267
        Top =0
        Name ="ChildPart_Cost_OEC_summed"
        Name =""
    End
    Begin
        Left =390
        Top =12
        Right =534
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
