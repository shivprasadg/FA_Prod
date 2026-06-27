Operation =1
Option =0
Begin InputTables
    Name ="ExitOptionsGroup_Local"
    Alias ="E"
    Name ="Units"
    Alias ="U"
    Name ="vw_BookValue"
    Alias ="B"
End
Begin OutputColumns
    Expression ="E.Client"
    Expression ="E.Schedule"
    Expression ="E.UnitGroup"
    Expression ="E.Assignee"
    Expression ="E.LeaseType"
    Expression ="B.BDEX"
    Expression ="E.BLCD"
    Expression ="E.AssetID"
    Expression ="E.UnitNum"
    Expression ="E.VIN"
    Alias ="Year"
    Expression ="U.UnitYr"
    Alias ="Make"
    Expression ="U.UnitMake"
    Alias ="Model"
    Expression ="U.UnitModel"
    Alias ="SubType"
    Expression ="U.UnitSubType"
    Alias ="Type"
    Expression ="U.UnitType"
    Expression ="E.CurrRent"
    Expression ="E.CurrRentType"
    Expression ="E.uStatus"
    Alias ="Primary Exp"
    Expression ="E.EOLPrimary"
    Expression ="E.DaysToEOL"
    Alias ="Exit Options"
    Expression ="E.Selected"
    Alias ="ExtdStart"
    Expression ="E.ExtendedStart"
    Alias ="ExtdEnd"
    Expression ="E.OffLeaseProj"
    Expression ="E.SufSent"
    Expression ="E.SufBack"
    Expression ="E.iPassedDate"
    Expression ="E.ActualOffLease"
    Expression ="B.LeaseRV"
    Expression ="B.BookNow"
    Alias ="Book 90d"
    Expression ="B.BV90Days"
    Expression ="E.ODORead"
    Expression ="E.ODOSource"
    Expression ="E.ODOReadDate"
    Expression ="E.Garage"
    Expression ="E.ExchDate"
    Expression ="E.DaysToExch"
    Expression ="E.NewClient"
    Expression ="E.NewSchedule"
    Expression ="E.NewGroup"
    Expression ="E.NewAssetID"
    Expression ="E.nInServ"
    Expression ="E.ReplacedSchNo"
    Expression ="E.ExchangeITMonth"
    Expression ="E.ExchangeITType"
    Expression ="E.OldSchGrp"
    Expression ="E.ClientID"
    Expression ="E.SchID"
    Expression ="E.GroupID"
    Expression ="E.MLOrig"
    Expression ="E.Selected"
End
Begin Joins
    LeftTable ="E"
    RightTable ="U"
    Expression ="E.AssetID = U.UnitID"
    Flag =1
    LeftTable ="E"
    RightTable ="B"
    Expression ="E.AssetID = B.AssetID"
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
        dbText "Name" ="E.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ODORead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Garage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.DaysToEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ODOSource"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ODOReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ReplacedSchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ActualOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.NewGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ExchDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.NewAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.DaysToExch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.NewClient"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.NewSchedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ExchangeITType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.OldSchGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Book 90d"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.SUFBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.SUFSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.uStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.nInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.CurrRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.CurrRentType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.iPassedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B.LeaseRV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Book@Expire"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtdEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtdStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Primary Exp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exit Options"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.EXITQTR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.EOLQTR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysToEXP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.EOLYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ExchTyp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.[Unit#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ExtendedStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.RentType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.SchIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ExtendedRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Sts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Selected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.LeaseExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ProjOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Book Expire"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.NewInServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ActiveRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjRV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.LastRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Mlorig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.EOLPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.OffLeaseProj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ExtRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Primary Expire"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Book @Expire"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exit Option"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B.BookNow"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B.BDEX"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =5
    Top =114
    Right =1242
    Bottom =1044
    Left =-1
    Top =-1
    Right =1213
    Bottom =398
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =265
        Bottom =453
        Top =0
        Name ="E"
        Name =""
    End
    Begin
        Left =359
        Top =9
        Right =503
        Bottom =153
        Top =0
        Name ="U"
        Name =""
    End
    Begin
        Left =376
        Top =184
        Right =520
        Bottom =445
        Top =0
        Name ="B"
        Name =""
    End
End
