Operation =1
Option =0
Where ="(((UnitExitOption.OptionTypeID)=3 Or (UnitExitOption.OptionTypeID) Is Null))"
Begin InputTables
    Name ="Units"
    Name ="UnitExitOption"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitYr"
    Expression ="UnitExitOption.OptionTypeID"
    Alias ="ActvRent"
    Expression ="CCur(Nz([unitrent],[Rent]))"
    Alias ="PriEOL"
    Expression ="DateAdd(\"m\",[LeaseTermPrimary],[BLCD])-1"
    Alias ="ExtendedEOL"
    Expression ="UnitExitOption.EndOrActualDate"
    Alias ="ProjEOL"
    Expression ="Nz([EndOrActualDate],[PriEOL])"
    Alias ="ActualEOL"
    Expression ="Units.unitoffleasedt"
    Alias ="EOLType"
    Expression ="IIf(IsNull([unitoffleasedt]),IIf(IsNull([EndOrActualDate]),\"Primary\",\"Extende"
        "d\"),\"Actual\")"
    Alias ="MonthsDiff"
    Expression ="IIf(IsNull([ActualEOL]),DateDiff(\"m\",[PriEOL],[ProjEOL]),DateDiff(\"m\",[PriEO"
        "L],[ActualEOL]))"
    Alias ="GAPType"
    Expression ="Switch([MonthsDiff]<0,\"Early:\" & Abs([MonthsDiff]),[MonthsDiff]=0,\"None\",[Mo"
        "nthsDiff]>0,\"GAP: \" & Abs([MonthsDiff]))"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="UnitExitOption"
    Expression ="Units.UnitID = UnitExitOption.UnitID"
    Flag =2
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =1
End
Begin OrderBy
    Expression ="Units.UnitID"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionTypeID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EOLType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PriEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActvRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtendedEOL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MonthsDiff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GAPType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =116
    Top =37
    Right =1435
    Bottom =1018
    Left =-1
    Top =-1
    Right =1295
    Bottom =629
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =21
        Top =12
        Right =192
        Bottom =404
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =302
        Top =81
        Right =517
        Bottom =447
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
    Begin
        Left =557
        Top =228
        Right =729
        Bottom =475
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
