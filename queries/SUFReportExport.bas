Operation =1
Option =0
Where ="((Not (E.SufSent) Is Null) AND ((E.ActualOffLease) Is Null))"
Begin InputTables
    Name ="ExitOptionsGroup_Local"
    Alias ="E"
    Name ="Units"
    Alias ="U"
    Name ="vw_BookValue"
    Alias ="B"
    Name ="vw_UnitInspectionLast"
    Name ="LastExtendedEvent"
End
Begin OutputColumns
    Expression ="E.Client"
    Expression ="E.Schedule"
    Expression ="E.UnitGroup"
    Expression ="E.Assignee"
    Expression ="E.LeaseType"
    Expression ="E.BLCD"
    Expression ="E.AssetID"
    Expression ="E.UnitNum"
    Expression ="E.VIN"
    Alias ="Year"
    Expression ="U.UnitYr"
    Alias ="Make"
    Expression ="E.UnitMake"
    Alias ="Model"
    Expression ="E.UnitModel"
    Alias ="SubType"
    Expression ="E.UnitSubType"
    Alias ="Type"
    Expression ="E.UnitType"
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
    Alias ="SentAge"
    Expression ="DateDiff(\"m\",[SufSent],Nz([SUFBack],Date()))"
    Alias ="BackAge"
    Expression ="DateDiff(\"m\",[SufBack],Date())"
    Expression ="vw_UnitInspectionLast.iDate"
    Expression ="E.iPassedDate"
    Expression ="E.ActualOffLease"
    Expression ="B.LeaseRV"
    Expression ="B.BookNow"
    Expression ="B.BV90Days"
    Expression ="E.ODORead"
    Expression ="E.ODOSource"
    Expression ="E.ODOReadDate"
    Expression ="E.Garage"
    Expression ="E.ExchDate"
    Expression ="E.NewClient"
    Expression ="E.NewSchedule"
    Expression ="E.NewGroup"
    Expression ="E.NewAssetID"
    Expression ="E.nInServ"
    Expression ="E.ClientID"
    Expression ="E.SchID"
    Expression ="E.GroupID"
    Expression ="LastExtendedEvent.LeaseOptionID"
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
    LeftTable ="E"
    RightTable ="vw_UnitInspectionLast"
    Expression ="E.AssetID = vw_UnitInspectionLast.AssetID"
    Flag =2
    LeftTable ="U"
    RightTable ="LastExtendedEvent"
    Expression ="U.UnitID = LastExtendedEvent.UnitID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[SUFReportExport].[Client]"
Begin
    Begin
        dbText "Name" ="SentAge"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BackAge"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.SufSent"
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
        dbText "Name" ="E.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.CurrRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
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
        dbText "Name" ="E.ODOSource"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.nInServ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ActualOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.uStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.CurrRentType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Primary Exp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.DaysToEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exit Options"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtdStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtdEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.SufBack"
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
        dbText "Name" ="B.BookNow"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B.BV90Days"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ODORead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ODOReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.Garage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ExchDate"
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
        dbText "Name" ="E.NewGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.NewAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UnitInspectionLast.iDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastExtendedEvent.LeaseOptionID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =175
    Top =54
    Right =2159
    Bottom =1129
    Left =-1
    Top =-1
    Right =1960
    Bottom =364
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =392
        Top =0
        Name ="E"
        Name =""
    End
    Begin
        Left =264
        Top =12
        Right =408
        Bottom =408
        Top =0
        Name ="U"
        Name =""
    End
    Begin
        Left =458
        Top =159
        Right =602
        Bottom =535
        Top =0
        Name ="B"
        Name =""
    End
    Begin
        Left =679
        Top =178
        Right =823
        Bottom =322
        Top =0
        Name ="vw_UnitInspectionLast"
        Name =""
    End
    Begin
        Left =795
        Top =42
        Right =939
        Bottom =186
        Top =0
        Name ="LastExtendedEvent"
        Name =""
    End
End
