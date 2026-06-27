Operation =1
Option =0
Begin InputTables
    Name ="RemarketingReportAllUnits"
End
Begin OutputColumns
    Expression ="RemarketingReportAllUnits.SourceCompany"
    Expression ="RemarketingReportAllUnits.UnitID"
    Expression ="RemarketingReportAllUnits.[Sch|Grp]"
    Expression ="RemarketingReportAllUnits.Unitnum"
    Expression ="RemarketingReportAllUnits.UnitYr"
    Expression ="RemarketingReportAllUnits.UnitMake"
    Expression ="RemarketingReportAllUnits.UnitModel"
    Expression ="RemarketingReportAllUnits.UnitType"
    Expression ="RemarketingReportAllUnits.UnitSubType"
    Expression ="RemarketingReportAllUnits.UnitStatus"
    Expression ="RemarketingReportAllUnits.InspectedDate"
    Expression ="RemarketingReportAllUnits.[SUF Sent]"
    Expression ="RemarketingReportAllUnits.[SUF Returned]"
    Expression ="RemarketingReportAllUnits.ParkDate"
    Expression ="RemarketingReportAllUnits.OffLeaseDate"
    Expression ="RemarketingReportAllUnits.BuyOutDate"
    Expression ="RemarketingReportAllUnits.InvoicePaid"
    Alias ="Parked To BuyOut"
    Expression ="IIf(IsDate([ParkDate]) And IsDate([BuyOutdate]) And ([ParkDate]<[BuyOutdate]),[B"
        "uyOutdate]-[ParkDate],Null)"
    Alias ="BuyOut to Sold"
    Expression ="IIf(IsDate([InvoicePaid]) And IsDate([BuyOutdate]) And [InvoicePaid]>=[BuyOutDat"
        "e],[InvoicePaid]-[Buyoutdate],IIf([InvoicePaid]<[BuyOutDate],0,Null))"
    Alias ="Parked To Sold"
    Expression ="IIf(IsDate([ParkDate]) And IsDate([InvoicePaid]),[InvoicePaid]-[Parkdate],Null)"
    Alias ="Bought Still in Inv"
    Expression ="IIf(Not IsDate([InvoicePaid]) And IsDate([BuyOutdate]),Date()-[BuyOutdate])"
    Alias ="Parked to OffLease"
    Expression ="IIf(IsDate([ParkDate]) And IsDate([OffLeaseDate]),[ParkDate]-[OffLeaseDate],Null"
        ")"
    Alias ="OffLease to Sold"
    Expression ="IIf(IsDate([InvoicePaid]) And IsDate([OffLeaseDate]),[InvoicePaid]-[OffLeaseDate"
        "],Null)"
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
        dbText "Name" ="RemarketingReportAllUnits.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.SourceCompany"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[Sch|Grp]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InspectedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.ParkDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.BuyOutDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoicePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.OffLeaseDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[SUF Sent]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.[SUF Returned]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parked to BuyOut"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyOut to Sold"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bought Still in Inv"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parked To Sold"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="OffLease to Park"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="OffLease to Sold"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketingReportAllUnits.InvoiceSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Still in Inv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parked to OffLease"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =182
    Top =11
    Right =1659
    Bottom =997
    Left =-1
    Top =-1
    Right =1459
    Bottom =709
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =61
        Top =22
        Right =205
        Bottom =344
        Top =0
        Name ="RemarketingReportAllUnits"
        Name =""
    End
End
