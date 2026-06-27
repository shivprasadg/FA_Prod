Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
    Name ="UnitLastMileageRead"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitStatus"
    Expression ="Units.UnitInServ"
    Alias ="ActualDelDate"
    Expression ="IIf([UnitActualDelDate]>Date(),Null,[UnitActualDelDate])"
    Expression ="Units.UnitAcceptDate"
    Alias ="StartDate"
    Expression ="CDate(Nz([UnitInServ],Nz([UnitAcceptDate],Nz([ActualDelDate],#1/1/1900#))))"
    Expression ="Rmkt.BuyOutDate"
    Alias ="LastReadDate"
    Expression ="IIf([unitstatus]=\"O\",Nz([Inspected2],Nz([Inspected2],Nz([ReadDate],Date()))),N"
        "ull)"
    Expression ="Units.UnitOLInd"
    Alias ="OffLeaseDate"
    Expression ="IIf([unitoffleasedt]>Date() Or [UnitOLInd]=\"P\",Null,[unitoffleasedt])"
    Expression ="Rmkt.ParkDT"
    Expression ="Rmkt.InvoicePaidDate"
    Alias ="EndDate"
    Expression ="CDate(Nz([OffLeaseDate],Nz([BuyOutDate],Nz([ParkDT],Nz([InvoicePaidDate],Nz([Las"
        "tReadDate],#12/31/2030#))))))"
    Expression ="Units.ReportStartDate"
    Expression ="Units.ReportEndDate"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =2
    LeftTable ="Units"
    RightTable ="UnitLastMileageRead"
    Expression ="Units.UnitID = UnitLastMileageRead.UnitID"
    Flag =2
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
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="StartDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="OffLeaseDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="EndDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Rmkt.ParkDT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Rmkt.BuyOutDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="ActualDelDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Units.ReportStartDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOLInd"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastReadDate"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1617
    Bottom =991
    Left =-1
    Top =-1
    Right =1585
    Bottom =438
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =76
        Top =86
        Right =241
        Bottom =366
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =292
        Top =123
        Right =436
        Bottom =267
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =508
        Top =158
        Right =652
        Bottom =302
        Top =0
        Name ="UnitLastMileageRead"
        Name =""
    End
End
