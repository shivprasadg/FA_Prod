Operation =1
Option =0
Where ="(((RemarketingReportAllUnits.UnitStatus)<>\"S\" And (RemarketingReportAllUnits.U"
    "nitStatus)<>\"C\" And (RemarketingReportAllUnits.UnitStatus)<>'SB' And (Remarket"
    "ingReportAllUnits.UnitStatus)<>\"KS\" And (RemarketingReportAllUnits.UnitStatus)"
    "<>\"D\"))"
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Expression ="RemarketingReportAllUnits.UnitID"
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Expression ="RemarketingReportAllUnits.[Sch|Grp]"
    Expression ="RemarketingReportAllUnits.UnitVIN"
    Expression ="RemarketingReportAllUnits.UnitDescShortVer"
    Expression ="RemarketingReportAllUnits.ODORead"
    Expression ="RemarketingReportAllUnits.ODOReadDate"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="RemarketingReportAllUnits.ParkDateProjected"
    Expression ="RemarketingReportAllUnits.OffLeaseDate"
    Expression ="RemarketingReportAllUnits.Residual"
    Expression ="RemarketingReportAllUnits.TermPenalty"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Expression ="RemarketingReportAllUnits.BuyOut"
    Expression ="RemarketingReportAllUnits.FMVBase"
    Expression ="RemarketingReportAllUnits.FMV"
    Expression ="RemarketingReportAllUnits.TargetSalePrice"
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
        dbText "Name" ="RemarketingReportAllUnits.SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ODORead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ODOReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[Sch|Grp]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Residual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitDescShortVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.OffLeaseDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ParkDateProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.FMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.FMVBase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.TargetSalePrice"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =111
    Top =94
    Right =1982
    Bottom =769
    Left =-1
    Top =-1
    Right =1847
    Bottom =152
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =358
        Top =27
        Right =502
        Bottom =566
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End
