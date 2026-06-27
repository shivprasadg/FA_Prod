Operation =1
Option =0
Where ="(((FleetStudyImportBase.VIN) Is Null))"
Begin InputTables
    Name ="FleetStudyImportBase"
    Name ="Non_Reporting_Units"
End
Begin OutputColumns
    Expression ="Non_Reporting_Units.*"
    Alias ="UnitX"
    Expression ="Right([UnitUnitNum],Len([UnitUnitNum])-1)"
End
Begin Joins
    LeftTable ="Non_Reporting_Units"
    RightTable ="FleetStudyImportBase"
    Expression ="Non_Reporting_Units.UnitVIN = FleetStudyImportBase.VIN"
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
        dbText "Name" ="Non_Reporting_Units.Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.OBCUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.Unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =30
    Right =1592
    Bottom =981
    Left =-1
    Top =-1
    Right =1517
    Bottom =633
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =626
        Top =102
        Right =770
        Bottom =246
        Top =0
        Name ="FleetStudyImportBase"
        Name =""
    End
    Begin
        Left =392
        Top =49
        Right =536
        Bottom =423
        Top =0
        Name ="Non_Reporting_Units"
        Name =""
    End
End
