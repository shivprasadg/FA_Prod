Operation =1
Option =0
Where ="(((Units.ClientGroupID)=[Ask ClientGroupID]) AND ((Units.UnitType)=\"daycab\" Or"
    " (Units.UnitType)=\"sleeper\" Or (Units.UnitType) Like \"*Straight*\") AND ((Uni"
    "ts.Unitstatus)=\"A\" Or (Units.Unitstatus)=\"T\" Or (Units.Unitstatus)=\"TR\") A"
    "ND ((Reporting_Units_CurrMonth.UnitID) Is Null))"
Begin InputTables
    Name ="Units"
    Name ="vw_SixKeys"
    Name ="Reporting_Units_CurrMonth"
End
Begin OutputColumns
    Expression ="Units.ClientGroupID"
    Expression ="Units.ClientID"
    Expression ="vw_SixKeys.CompanyName"
    Expression ="vw_SixKeys.MLNo"
    Expression ="Units.UnitID"
    Expression ="Units.UnitVIN"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.OBCUnitNumber"
    Expression ="Units.UnitYr"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitType"
    Expression ="Units.UnitSubType"
    Expression ="Units.Unitstatus"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="vw_SixKeys"
    Expression ="Units.UnitID = vw_SixKeys.UnitID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Reporting_Units_CurrMonth"
    Expression ="Units.UnitID = Reporting_Units_CurrMonth.UnitID"
    Flag =2
End
Begin OrderBy
    Expression ="Units.ClientID"
    Flag =0
    Expression ="Units.UnitID"
    Flag =0
    Expression ="Units.OBCUnitNumber"
    Flag =0
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
dbMemo "OrderBy" ="[Non_Reporting_Units].[UnitYr] DESC"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.OBCUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.Unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =108
    Top =52
    Right =1657
    Bottom =1003
    Left =-1
    Top =-1
    Right =1517
    Bottom =599
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =299
        Top =49
        Right =443
        Bottom =520
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =694
        Top =263
        Right =844
        Bottom =600
        Top =0
        Name ="vw_Client_FAID_PO_Lease_Sch_Grp_Units"
        Name =""
    End
    Begin
        Left =571
        Top =83
        Right =799
        Bottom =227
        Top =0
        Name ="Reporting_Units_CurrMonth"
        Name =""
    End
End
