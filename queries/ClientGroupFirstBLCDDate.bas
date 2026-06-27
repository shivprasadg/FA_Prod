Operation =1
Option =0
Where ="(((ProposalsSummaryQuery.dealOutcome)=\"Awarded\"))"
Having ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="ProposalsSummaryQuery"
    Name ="Clients"
End
Begin OutputColumns
    Alias ="AwardedDate"
    Expression ="Min(Clients.ClientBecamedate)"
    Alias ="MinOfBLCD"
    Expression ="Min(vw_SixKeys.BLCD)"
    Alias ="ClientGroupName"
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.AssetID)"
    Alias ="2009"
    Expression ="IIf(AwardedDate<#2/1/2009#,[CLientGroupName],Null)"
    Alias ="2009 Units"
    Expression ="IIf(AwardedDate<#2/1/2009#,Count([AssetID]),Null)"
    Alias ="2010"
    Expression ="IIf(AwardedDate Between #2/1/2009# And #2/1/2011#,[CLientGroupName],Null)"
    Alias ="2010 Units"
    Expression ="IIf(AwardedDate Between #2/1/2009# And #2/1/2011#,Count([AssetId]),Null)"
    Alias ="2011"
    Expression ="IIf(AwardedDate Between #2/1/2010# And #2/1/2012#,[CLientGroupName],Null)"
    Alias ="2011 Units"
    Expression ="IIf(AwardedDate Between #2/1/2011# And #2/1/2012#,Count([AssetId]),Null)"
    Alias ="2012"
    Expression ="IIf(AwardedDate Between #2/1/2012# And #2/1/2013#,[CLientGroupName],Null)"
    Alias ="2012 Units"
    Expression ="IIf(AwardedDate Between #2/1/2012# And #2/1/2013#,Count([AssetId]),Null)"
    Alias ="2013"
    Expression ="IIf(AwardedDate Between #2/1/2013# And #2/1/2014#,[CLientGroupName],Null)"
    Alias ="2013 Units"
    Expression ="IIf(AwardedDate Between #2/1/2013# And #2/1/2014#,Count([AssetId]),Null)"
    Alias ="2014"
    Expression ="IIf(AwardedDate Between #2/1/2014# And #2/1/2015#,[CLientGroupName],Null)"
    Alias ="2014 Units"
    Expression ="IIf(AwardedDate Between #2/1/2014# And #2/1/2015#,Count([AssetId]),Null)"
    Alias ="2015"
    Expression ="IIf(AwardedDate Between #2/1/2015# And #2/1/2016#,[CLientGroupName],Null)"
    Alias ="2015 Units"
    Expression ="IIf(AwardedDate Between #2/1/2015# And #2/1/2016#,Count([AssetId]),Null)"
    Alias ="2016"
    Expression ="IIf(AwardedDate Between #2/1/2016# And #2/1/2017#,[CLientGroupName],Null)"
    Alias ="2016 Units"
    Expression ="IIf(AwardedDate Between #2/1/2016# And #2/1/2017#,Count([AssetId]),Null)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="ProposalsSummaryQuery"
    Expression ="vw_SixKeys.clientID = ProposalsSummaryQuery.clientRef"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Clients"
    Expression ="vw_SixKeys.clientID = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="Min(Clients.ClientBecamedate)"
    Flag =1
End
Begin Groups
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
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
dbBoolean "TotalsRow" ="-1"
Begin
    Begin
        dbText "Name" ="2009 Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="2010 Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MinOfBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2009"
        dbLong "AggregateType" ="2"
    End
    Begin
        dbText "Name" ="2010"
        dbLong "AggregateType" ="2"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="2011"
        dbLong "AggregateType" ="2"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="2011 Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="2012"
        dbLong "AggregateType" ="2"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="2012 Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="2013 Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="2013"
        dbLong "AggregateType" ="2"
    End
    Begin
        dbText "Name" ="2015 Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2014"
        dbLong "AggregateType" ="2"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="2015"
        dbLong "AggregateType" ="2"
    End
    Begin
        dbText "Name" ="2014 Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2016"
        dbLong "AggregateType" ="2"
    End
    Begin
        dbText "Name" ="2016 Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AwardedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientGroupName"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-1
    Top =22
    Right =1898
    Bottom =903
    Left =-1
    Top =-1
    Right =1867
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =381
        Top =55
        Right =659
        Bottom =394
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =732
        Top =119
        Right =1050
        Bottom =443
        Top =0
        Name ="ProposalsSummaryQuery"
        Name =""
    End
    Begin
        Left =-9
        Top =21
        Right =348
        Bottom =410
        Top =0
        Name ="Clients"
        Name =""
    End
End
