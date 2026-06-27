Operation =1
Option =0
Where ="(((vw_SixKeys.Status)=\"A\" Or (vw_SixKeys.Status)=\"TR\") AND ((vw_SixKeys.MLOr"
    "ig)=\"FA\") AND ((vw_Units_AutoHaulerPaired.TractorUnitID) Is Null) AND ((vw_Uni"
    "ts_AutoHaulerPaired_1.AutoHaulerUnitID) Is Null))"
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
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
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
    Expression ="Sum(UnitTotalOEC.UnitTotalCost)"
    Alias ="RV"
    Expression ="CCur(Nz([Units].[UnitSynResid],[Units].[UnitSalesResid]))"
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
    Flag =2
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
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    GroupLevel =0
    Expression ="vw_SixKeys.LeaseTermPrimary"
    GroupLevel =0
    Expression ="vw_SixKeys.ProjEOL"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITMonth"
    GroupLevel =0
    Expression ="CCur(Nz([Units].[UnitSynResid],[Units].[UnitSalesResid]))"
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
dbBoolean "TotalsRow" ="-1"
dbMemo "Filter" ="(Year([ExpiryReport_Base].[ExchangeITDate])=Year(Date()))"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.BDEX"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExpiryDate"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoticeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITTerm"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Term"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbInteger "ColumnWidth" ="3120"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="RV"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-31
    Right =128
    Bottom =8
    Left =-1
    Top =-1
    Right =1404
    Bottom =615
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =44
        Top =21
        Right =341
        Bottom =493
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =388
        Top =74
        Right =532
        Bottom =630
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =760
        Top =165
        Right =904
        Bottom =461
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
    Begin
        Left =580
        Top =12
        Right =724
        Bottom =698
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =772
        Top =12
        Right =916
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =964
        Top =12
        Right =1108
        Bottom =156
        Top =0
        Name ="vw_Units_AutoHaulerPaired"
        Name =""
    End
    Begin
        Left =964
        Top =173
        Right =1108
        Bottom =317
        Top =0
        Name ="vw_Units_AutoHaulerPaired_1"
        Name =""
    End
End
