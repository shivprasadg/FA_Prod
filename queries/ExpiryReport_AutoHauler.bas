Operation =1
Option =0
Where ="(((vw_SixKeys.Status)=\"A\" Or (vw_SixKeys.Status)=\"TR\") AND ((vw_SixKeys.MLOr"
    "ig)=\"FA\") AND ((vw_Units_AutoHaulerPaired_1.AutoHaulerUnitID) Is Null))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="SchGrp"
    Name ="UnitTotalOEC"
    Name ="Locations"
    Name ="Units"
    Name ="vw_Units_AutoHaulerPaired"
    Name ="vw_Units_AutoHaulerPaired"
    Alias ="vw_Units_AutoHaulerPaired_1"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.BDEX"
    Expression ="vw_SixKeys.Schedule"
    Alias ="UnitType"
    Expression ="[vw_Units_AutoHaulerPaired].[TractorSubType] & \"-AHauler\""
    Alias ="Units"
    Expression ="Count(vw_SixKeys.UnitID)"
    Alias ="Term"
    Expression ="vw_SixKeys.LeaseTermPrimary"
    Alias ="ExpiryDate"
    Expression ="vw_SixKeys.ProjEOL"
    Alias ="ExchangeITTerm"
    Expression ="SchGrp.ExchangeITMonth"
    Alias ="ExchangeITDate"
    Expression ="Min(DateAdd(\"m\",[ExchangeITMonth],[SchGrp].[BLCD]))"
    Alias ="NoticeDate"
    Expression ="Min(DateAdd(\"m\",-3,[ProjEOL]))"
    Alias ="OEC"
    Expression ="vw_Units_AutoHaulerPaired.PairedCost"
    Alias ="RV"
    Expression ="CCur([vw_Units_AutoHaulerPaired].[PairedRV])"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="UnitTotalOEC"
    Expression ="vw_SixKeys.UnitID = UnitTotalOEC.AssetID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Locations"
    Expression ="vw_SixKeys.LocationID = Locations.LocationID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="vw_Units_AutoHaulerPaired"
    Expression ="vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired.TractorUnitID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="vw_Units_AutoHaulerPaired_1"
    Expression ="vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired_1.AutoHaulerUnitID"
    Flag =2
End
Begin OrderBy
    Expression ="vw_SixKeys.BDEX"
    Flag =0
    Expression ="vw_SixKeys.CompanyName"
    Flag =0
    Expression ="vw_SixKeys.Schedule"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.BDEX"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="[vw_Units_AutoHaulerPaired].[TractorSubType] & \"-AHauler\""
    GroupLevel =0
    Expression ="vw_SixKeys.LeaseTermPrimary"
    GroupLevel =0
    Expression ="vw_SixKeys.ProjEOL"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITMonth"
    GroupLevel =0
    Expression ="vw_Units_AutoHaulerPaired.PairedCost"
    GroupLevel =0
    Expression ="CCur([vw_Units_AutoHaulerPaired].[PairedRV])"
    GroupLevel =0
    Expression ="vw_SixKeys.CompanyName"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
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
Begin
    Begin
        dbText "Name" ="Term"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUnitOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="ExpiryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoticeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Units_AutoHaulerPaired_1.TractorYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Units_AutoHaulerPaired_1.TractorVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Units_AutoHaulerPaired.UnitOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Units_AutoHaulerPaired.PairedCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Units_AutoHaulerPaired.PairedRV"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =171
    Top =0
    Right =1373
    Bottom =977
    Left =-1
    Top =-1
    Right =1184
    Bottom =632
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =634
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =495
        Top =264
        Right =639
        Bottom =636
        Top =0
        Name ="vw_Units_AutoHaulerPaired"
        Name =""
    End
    Begin
        Left =846
        Top =164
        Right =1068
        Bottom =633
        Top =0
        Name ="vw_Units_AutoHaulerPaired_1"
        Name =""
    End
End
