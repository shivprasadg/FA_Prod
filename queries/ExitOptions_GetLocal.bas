Operation =3
Name ="ExitOptionsGroup_Local"
Option =0
Begin InputTables
    Name ="ExitOptionsGroup_EOL"
End
Begin OutputColumns
    Name ="AssetID"
    Expression ="ExitOptionsGroup_EOL.AssetID"
    Name ="Unit#"
    Expression ="ExitOptionsGroup_EOL.[Unit#]"
    Name ="VIN"
    Expression ="ExitOptionsGroup_EOL.VIN"
    Name ="Make"
    Expression ="ExitOptionsGroup_EOL.Make"
    Name ="ClientShortName"
    Expression ="ExitOptionsGroup_EOL.ClientShortName"
    Name ="Schedule"
    Expression ="ExitOptionsGroup_EOL.Schedule"
    Name ="UnitGroup"
    Expression ="ExitOptionsGroup_EOL.UnitGroup"
    Name ="ExchTyp"
    Expression ="ExitOptionsGroup_EOL.ExchTyp"
    Name ="Sts"
    Expression ="ExitOptionsGroup_EOL.Sts"
    Name ="Selected"
    Expression ="ExitOptionsGroup_EOL.Selected"
    Name ="EOLDate"
    Expression ="ExitOptionsGroup_EOL.EOLDate"
    Name ="EOLDatePrimary"
    Expression ="ExitOptionsGroup_EOL.EOLDatePrimary"
    Name ="DaysToEOLPrimary"
    Expression ="ExitOptionsGroup_EOL.DaysToEOLPrimary"
    Name ="DaysToEOL"
    Expression ="ExitOptionsGroup_EOL.DaysToEOL"
    Name ="ExchDate"
    Expression ="ExitOptionsGroup_EOL.ExchDate"
    Name ="DaysToExch"
    Expression ="ExitOptionsGroup_EOL.DaysToExch"
    Name ="Assignee"
    Expression ="ExitOptionsGroup_EOL.Assignee"
    Name ="ActualOffLeaseDate"
    Expression ="ExitOptionsGroup_EOL.ActualOffLeaseDate"
    Name ="OffLeaseInd"
    Expression ="ExitOptionsGroup_EOL.OffLeaseInd"
    Name ="ODORead"
    Expression ="ExitOptionsGroup_EOL.ODORead"
    Name ="ODOSource"
    Expression ="ExitOptionsGroup_EOL.ODOSource"
    Name ="ODOReadDate"
    Expression ="ExitOptionsGroup_EOL.ODOReadDate"
    Name ="Garage"
    Expression ="ExitOptionsGroup_EOL.Garage"
    Name ="LastModified"
    Expression ="ExitOptionsGroup_EOL.LastModified"
    Name ="NewClient"
    Expression ="ExitOptionsGroup_EOL.NewClient"
    Name ="NewSchedule"
    Expression ="ExitOptionsGroup_EOL.NewSchedule"
    Name ="NewGroup"
    Expression ="ExitOptionsGroup_EOL.NewGroup"
    Name ="NewAssetID"
    Expression ="ExitOptionsGroup_EOL.NewAssetID"
    Name ="NewInServiceDate"
    Expression ="ExitOptionsGroup_EOL.NewInServiceDate"
    Name ="MTMRent"
    Expression ="ExitOptionsGroup_EOL.MTMRent"
    Name ="ExtendedRent"
    Expression ="ExitOptionsGroup_EOL.ExtendedRent"
    Name ="ActiveRent"
    Expression ="ExitOptionsGroup_EOL.ActiveRent"
    Name ="RentType"
    Expression ="ExitOptionsGroup_EOL.RentType"
    Name ="ExtendedStart"
    Expression ="ExitOptionsGroup_EOL.ExtendedStart"
    Name ="ExtendedEOL"
    Expression ="ExitOptionsGroup_EOL.ExtendedEOL"
    Name ="ReplacedGroupNo"
    Expression ="ExitOptionsGroup_EOL.ReplacedGroupNo"
    Name ="ReplacedSchNo"
    Expression ="ExitOptionsGroup_EOL.ReplacedSchNo"
    Name ="ExchangeITMonth"
    Expression ="ExitOptionsGroup_EOL.ExchangeITMonth"
    Name ="ExchangeITType"
    Expression ="ExitOptionsGroup_EOL.ExchangeITType"
    Name ="ProjRV"
    Expression ="ExitOptionsGroup_EOL.ProjRV"
    Name ="LastRent"
    Expression ="ExitOptionsGroup_EOL.LastRent"
    Name ="OldSchGrp"
    Expression ="ExitOptionsGroup_EOL.OldSchGrp"
    Name ="OldAssetID"
    Expression ="ExitOptionsGroup_EOL.OldAssetID"
    Name ="InspectionLinkWeb"
    Expression ="ExitOptionsGroup_EOL.InspectionLinkWeb"
    Name ="InspectionLinkLocal"
    Expression ="ExitOptionsGroup_EOL.InspectionLinkLocal"
    Name ="SchIn"
    Expression ="ExitOptionsGroup_EOL.SchIn"
    Name ="EOLQTR"
    Expression ="ExitOptionsGroup_EOL.EOLQTR"
    Name ="EOLMth"
    Expression ="ExitOptionsGroup_EOL.EOLMth"
    Name ="EOLDay"
    Expression ="ExitOptionsGroup_EOL.EOLDay"
    Name ="EOLYear"
    Expression ="ExitOptionsGroup_EOL.EOLYear"
    Name ="EXITQTR"
    Expression ="ExitOptionsGroup_EOL.EXITQTR"
    Name ="EXITMth"
    Expression ="ExitOptionsGroup_EOL.EXITMth"
    Name ="EXITDay"
    Expression ="ExitOptionsGroup_EOL.EXITDay"
    Name ="EXITYear"
    Expression ="ExitOptionsGroup_EOL.EXITYear"
    Name ="ClientID"
    Expression ="ExitOptionsGroup_EOL.ClientID"
    Name ="SchID"
    Expression ="ExitOptionsGroup_EOL.SchID"
    Name ="GroupID"
    Expression ="ExitOptionsGroup_EOL.GroupID"
    Name ="BankId"
    Expression ="ExitOptionsGroup_EOL.BankId"
    Name ="GarageLocationID"
    Expression ="ExitOptionsGroup_EOL.GarageLocationID"
    Name ="ReplacedByScheduleID"
    Expression ="ExitOptionsGroup_EOL.ReplacedByScheduleID"
    Name ="ReplacedByGroupID"
    Expression ="ExitOptionsGroup_EOL.ReplacedByGroupID"
    Name ="MaxLeaseOptionID"
    Expression ="ExitOptionsGroup_EOL.MaxLeaseOptionID"
    Name ="ReplacedSchID"
    Expression ="ExitOptionsGroup_EOL.ReplacedSchID"
    Name ="ReplacedGroupID"
    Expression ="ExitOptionsGroup_EOL.ReplacedGroupID"
    Name ="OldSchID"
    Expression ="ExitOptionsGroup_EOL.OldSchID"
    Name ="OldGroupID"
    Expression ="ExitOptionsGroup_EOL.OldGroupID"
    Alias ="WorkDate"
    Name ="WorkDate"
    Expression ="Now()"
    Name ="LeaseExp"
    Expression ="ExitOptionsGroup_EOL.LeaseExp"
    Name ="ActualOffLease"
    Expression ="ExitOptionsGroup_EOL.ActualOffLease"
    Name ="ProjOffLease"
    Expression ="ExitOptionsGroup_EOL.ProjOffLease"
    Name ="MLOrig"
    Expression ="ExitOptionsGroup_EOL.MLOrig"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.OldGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WorkDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.OldSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.ReplacedGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.ProjOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.ActualOffLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.LeaseExp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExitOptionsGroup_EOL.MLOrig"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-462
    Top =20
    Right =862
    Bottom =949
    Left =-1
    Top =-1
    Right =1306
    Bottom =601
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =90
        Top =23
        Right =234
        Bottom =492
        Top =0
        Name ="ExitOptionsGroup_EOL"
        Name =""
    End
End
