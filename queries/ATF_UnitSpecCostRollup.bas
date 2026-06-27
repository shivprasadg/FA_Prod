Operation =1
Option =0
Begin InputTables
    Name ="ATFUnitSpec"
    Name ="ATF_ChildPartCost"
End
Begin OutputColumns
    Expression ="ATFUnitSpec.ATFID"
    Expression ="ATFUnitSpec.ATFGroupSpecID"
    Alias ="NumUnitsInSpec"
    Expression ="Sum(ATFUnitSpec.NumUnitsInSpec)"
    Alias ="UnitCost"
    Expression ="Sum(ATFUnitSpec.UnitCost)"
    Alias ="PartCost"
    Expression ="Sum(CCur(Nz([ChildPartCost],0)))"
    Alias ="TotalUnitCost"
    Expression ="[UnitCost]+CCur(Nz(Sum([ChildPartCost]),0))"
    Alias ="SchRent"
    Expression ="Sum(ATFUnitSpec.SchRent)"
    Alias ="LRFSch"
    Expression ="IIf(Nz([SchRent],0)=0 Or Nz([TotalUnitCost],0)=0,Null,Round([SchRent]/[TotalUnit"
        "Cost],5))"
    Alias ="FirstExtendedRent"
    Expression ="Sum(ATFUnitSpec.FirstExtendedRent)"
    Alias ="LRFExt"
    Expression ="IIf(Nz([FirstExtendedRent],0)=0 Or Nz([TotalUnitCost],0)=0,Null,Round([FirstExte"
        "ndedRent]/[TotalUnitCost],5))"
    Alias ="MTMRent"
    Expression ="Sum(ATFUnitSpec.MTMRent)"
    Alias ="LRFMTM"
    Expression ="IIf(Nz([MTMRent],0)=0 Or Nz([TotalUnitCost],0)=0,Null,Round([MTMRent]/[TotalUnit"
        "Cost],5))"
    Alias ="PerDiemRent"
    Expression ="Sum(ATFUnitSpec.PerDiemRent)"
    Alias ="FirstExtendedTerm"
    Expression ="Avg(ATFUnitSpec.FirstExtendedTerm)"
    Alias ="SalesResidual"
    Expression ="Sum(ATFUnitSpec.SalesResidual)"
    Alias ="SpecsIn"
    Expression ="Count(ATFUnitSpec.ATFUnitSpecID)"
    Alias ="UTARebate"
    Expression ="Sum(CCur(Nz([ATFUnitSpec].[UTARebate],0)))"
End
Begin Joins
    LeftTable ="ATFUnitSpec"
    RightTable ="ATF_ChildPartCost"
    Expression ="ATFUnitSpec.ATFUnitSpecID = ATF_ChildPartCost.ATFUnitSpecID"
    Flag =2
End
Begin Groups
    Expression ="ATFUnitSpec.ATFID"
    GroupLevel =0
    Expression ="ATFUnitSpec.ATFGroupSpecID"
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
        dbText "Name" ="ATFUnitSpec.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstExtendedRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiemRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalUnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LRFSch"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LRFExt"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LRFMTM"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpecsIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstExtendedTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTARebate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumUnitsInSpec"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =81
    Top =33
    Right =1592
    Bottom =856
    Left =-1
    Top =-1
    Right =1493
    Bottom =427
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =251
        Top =50
        Right =556
        Bottom =281
        Top =0
        Name ="ATFUnitSpec"
        Name =""
    End
    Begin
        Left =639
        Top =143
        Right =783
        Bottom =287
        Top =0
        Name ="ATF_ChildPartCost"
        Name =""
    End
End
