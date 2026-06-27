Operation =1
Option =0
Having ="(((vw_SixKeys.MLOrig)=\"FA\") AND ((Units.UnitsRent)>0) AND ((Units.UnitRent)>0)"
    ")"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Units"
End
Begin OutputColumns
    Expression ="vw_SixKeys.MLOrig"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.FAID"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Expression ="vw_SixKeys.UnitDescShortVer"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.assetID)"
    Alias ="SchRent"
    Expression ="Units.UnitsRent"
    Alias ="SyndRent"
    Expression ="Units.UnitRent"
    Alias ="SchVsSynd"
    Expression ="IIf([Units].[UnitRent]<[Units].[UnitSRent],\">SchRent\",'')"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="vw_SixKeys.MLID"
    Flag =0
    Expression ="vw_SixKeys.Schedule"
    Flag =0
    Expression ="vw_SixKeys.UnitGroup"
    Flag =0
    Expression ="vw_SixKeys.SchID"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.MLID"
    GroupLevel =0
    Expression ="vw_SixKeys.FAID"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitGroup"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitDescShortVer"
    GroupLevel =0
    Expression ="Units.UnitsRent"
    GroupLevel =0
    Expression ="Units.UnitRent"
    GroupLevel =0
    Expression ="IIf([Units].[UnitRent]<[Units].[UnitSRent],\">SchRent\",'')"
    GroupLevel =0
    Expression ="vw_SixKeys.SchID"
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
dbMemo "Filter" ="([FARents_ML_SchGrp_Desc].[SchVsSynd]=\">SchRent\")"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescShortVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchVsSynd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =109
    Top =17
    Right =1613
    Bottom =1021
    Left =-1
    Top =-1
    Right =1472
    Bottom =697
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =17
        Top =36
        Right =672
        Bottom =471
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =744
        Top =159
        Right =888
        Bottom =303
        Top =0
        Name ="Units"
        Name =""
    End
End
