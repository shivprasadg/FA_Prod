Operation =1
Option =0
Where ="(((Units.UnitPDStartDate) Is Not Null))"
Begin InputTables
    Name ="Units"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Alias ="GroupID"
    Expression ="Units.GroupREF"
    Alias ="PerdiemStartDate"
    Expression ="Units.UnitPDStartDate"
    Alias ="PerdiemRate"
    Expression ="Nz([UNITPDRENT],0)"
    Alias ="PerdiemDays"
    Expression ="DateDiff(\"d\",[UnitPDStartDate],Nz([BLCD],[BLCD_Projected])-1)"
    Alias ="PerdiemTotal"
    Expression ="IIf(Nz([PerdiemDays],0)<1,0,[PerdiemDays]*[UnitPDRent])"
    Alias ="PDType"
    Expression ="IIf(IsNull([schgrp].[BLCD]),\"Projected\",\"Actual\")"
    Expression ="SchGrp.FundDate_Projected"
    Expression ="SchGrp.BLCD"
    Alias ="FundDate"
    Expression ="IIf(IsNull([schgrp].[FundDate]),[schgrp].[FundDate_Projected],[schgrp].[FundDate"
        "])"
    Expression ="SchGrp.BLCD_Projected"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
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
dbMemo "Filter" ="([UnitGroupPerdiemsSyndReport].[UnitID]=323678)"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerdiemDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerdiemTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerdiemStartDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="PerdiemRate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =322
    Top =3
    Right =1635
    Bottom =886
    Left =-1
    Top =-1
    Right =1295
    Bottom =549
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =104
        Top =93
        Right =248
        Bottom =432
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =303
        Top =80
        Right =447
        Bottom =480
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
