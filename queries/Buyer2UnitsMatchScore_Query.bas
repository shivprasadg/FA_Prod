Operation =1
Option =0
Begin InputTables
    Name ="RemarketingUnitToBuyerScoreMatched"
End
Begin OutputColumns
    Alias ="SCR"
    Expression ="RemarketingUnitToBuyerScoreMatched.BuyerScore"
    Alias ="LastBuy"
    Expression ="RemarketingUnitToBuyerScoreMatched.TypePurchased"
    Alias ="Days"
    Expression ="RemarketingUnitToBuyerScoreMatched.DaysAgo"
    Alias ="Miles"
    Expression ="RemarketingUnitToBuyerScoreMatched.MilesFromBuyer"
    Expression ="RemarketingUnitToBuyerScoreMatched.UnitID"
    Alias ="Sts"
    Expression ="RemarketingUnitToBuyerScoreMatched.UnitStatus"
    Expression ="RemarketingUnitToBuyerScoreMatched.UnitType"
    Alias ="Year"
    Expression ="RemarketingUnitToBuyerScoreMatched.UnitYr"
    Alias ="UnitDesc"
    Expression ="UnitMake & ' ' & UnitModel & ' ' & Axle"
    Alias ="Price"
    Expression ="Format(Nz(TargetSale,0),'$#,###')"
    Alias ="Proj'd Park"
    Expression ="RemarketingUnitToBuyerScoreMatched.ParkDateProjected"
    Alias ="Parked City"
    Expression ="RemarketingUnitToBuyerScoreMatched.LCity"
    Alias ="St"
    Expression ="RemarketingUnitToBuyerScoreMatched.LState"
    Expression ="RemarketingUnitToBuyerScoreMatched.UnitVin"
    Expression ="RemarketingUnitToBuyerScoreMatched.ODORead"
    Alias ="ODODate"
    Expression ="Format(ODOReadDate,'mm/dd/yy')"
    Expression ="RemarketingUnitToBuyerScoreMatched.ODOSource"
    Expression ="RemarketingUnitToBuyerScoreMatched.Latitude"
    Expression ="RemarketingUnitToBuyerScoreMatched.Longitude"
    Expression ="RemarketingUnitToBuyerScoreMatched.HomeLat"
    Expression ="RemarketingUnitToBuyerScoreMatched.HomeLog"
    Expression ="RemarketingUnitToBuyerScoreMatched.clientID"
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
        dbText "Name" ="SCR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastBuy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Days"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Miles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj'd Park"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parked City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ODODate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1637
    Bottom =991
    Left =-1
    Top =-1
    Right =1605
    Bottom =463
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =336
        Bottom =378
        Top =0
        Name ="RemarketingUnitToBuyerScoreMatched"
        Name =""
    End
End
