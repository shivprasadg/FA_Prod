Operation =1
Option =0
Having ="(((vw_SixKeys.MLOrig)=\"FA\") AND ((IIf(DateDiff(\"m\",[ProjEOL],Nz([unitoffleas"
    "edt],Date()))>=1,1,0))=0) AND ((vw_SixKeys.OffLeaseDate) Between #1/1/2016# And "
    "#12/31/2016#))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
End
Begin OutputColumns
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.UnitID"
    Alias ="OnLease2015"
    Expression ="IIf([vw_SixKeys].[BLCD]<#1/1/2015# And Nz([unitoffleasedt],#1/1/2030#)>#12/31/20"
        "15#,1,0)"
    Alias ="Added2015"
    Expression ="IIf([BLCD] Between #1/1/2015# And #12/31/2015#,1,0)"
    Alias ="Terminated2015"
    Expression ="IIf([unitoffleasedt] Between #1/1/2015# And #12/31/2015#,1,0)"
    Alias ="Added2016"
    Expression ="IIf([BLCD] Between #1/1/2016# And #12/31/2016#,1,0)"
    Alias ="Terminated2016"
    Expression ="IIf([unitoffleasedt] Between #1/1/2016# And #12/31/2016#,1,0)"
    Alias ="Extended"
    Expression ="IIf(DateDiff(\"m\",[ProjEOL],Nz([unitoffleasedt],Date()))>=1,1,0)"
    Alias ="Extended2015"
    Expression ="IIf([OnLease2015]=1 And [Extended]=1,1,0)"
    Alias ="Extended2016"
    Expression ="IIf([Added2016]=1 And [Extended]=1,1,0)"
    Alias ="MonthsExt"
    Expression ="DateDiff(\"m\",[ProjEOL],Nz([unitoffleasedt],Date()))"
    Expression ="vw_SixKeys.ProjEOL"
    Expression ="vw_SixKeys.OffLeaseDate"
    Alias ="CalcEOL"
    Expression ="DateAdd(\"m\",[LeaseTermPrimary],[BLCD]-1)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitID"
    GroupLevel =0
    Expression ="IIf([vw_SixKeys].[BLCD]<#1/1/2015# And Nz([unitoffleasedt],#1/1/2030#)>#12/31/20"
        "15#,1,0)"
    GroupLevel =0
    Expression ="IIf([BLCD] Between #1/1/2015# And #12/31/2015#,1,0)"
    GroupLevel =0
    Expression ="IIf([unitoffleasedt] Between #1/1/2015# And #12/31/2015#,1,0)"
    GroupLevel =0
    Expression ="IIf([BLCD] Between #1/1/2016# And #12/31/2016#,1,0)"
    GroupLevel =0
    Expression ="IIf([unitoffleasedt] Between #1/1/2016# And #12/31/2016#,1,0)"
    GroupLevel =0
    Expression ="IIf(DateDiff(\"m\",[ProjEOL],Nz([unitoffleasedt],Date()))>=1,1,0)"
    GroupLevel =0
    Expression ="DateDiff(\"m\",[ProjEOL],Nz([unitoffleasedt],Date()))"
    GroupLevel =0
    Expression ="vw_SixKeys.ProjEOL"
    GroupLevel =0
    Expression ="vw_SixKeys.OffLeaseDate"
    GroupLevel =0
    Expression ="DateAdd(\"m\",[LeaseTermPrimary],[BLCD]-1)"
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
dbMemo "Filter" ="([LeaseInOutExtendedBase].[ClientShortName]=\"DYNO\")"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Added2015"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Terminated2016"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Terminated2015"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Added2016"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Extended"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnLease2015"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Extended2015"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Extended2016"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MonthsExt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.OffLeaseDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ProjEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CalcEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysExt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =105
    Top =302
    Right =1311
    Bottom =965
    Left =-1
    Top =-1
    Right =1468
    Bottom =377
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =285
        Bottom =422
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =333
        Top =97
        Right =477
        Bottom =241
        Top =0
        Name ="Units"
        Name =""
    End
End
