Operation =1
Option =0
Where ="(((FleetStudyImportBase.UnitNumber)<>Non_Reporting_Units.UnitUnitNum) And ((Flee"
    "tStudyImportBase.VIN) Is Not Null))"
Begin InputTables
    Name ="Non_Reporting_Units"
    Name ="FleetStudyImportBase"
End
Begin OutputColumns
    Expression ="Non_Reporting_Units.*"
    Alias ="UNFI_UnitNum"
    Expression ="FleetStudyImportBase.UnitNumber"
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
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.ClientGroupID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.ClientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.OBCUnitNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitYr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitSubType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.Unitstatus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.vw_Client_FAID_PO_Lease_Sch_Grp_Units.CompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Non_Reporting_Units.Units.UnitVIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="UNFI_UnitNum"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =245
    Top =-7
    Right =1794
    Bottom =944
    Left =-1
    Top =-1
    Right =1517
    Bottom =599
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
