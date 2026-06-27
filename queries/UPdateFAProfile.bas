Operation =4
Option =0
Begin InputTables
    Name ="clients"
End
Begin OutputColumns
    Name ="clients.fleetSize"
    Expression ="[updateFAProfile].[fleetDaycabs]+[updateFAProfile].[fleetSleepers]"
    Name ="clients.creditFAScore"
    Expression ="[updateFAProfile].[ScoreCredit]+[updateFAProfile].[ScoreCredit2]+[updateFAProfil"
        "e].[ScoreFleetSize]+[updateFAProfile].[ScorePercentDC]+[updateFAProfile].[ScoreD"
        "CMiles]+[updateFAProfile].[ScoreSLMiles]+[updateFAProfile].[ScoreMaintenance]+[u"
        "pdateFAProfile].[ScoreLifecycle]+[updateFAProfile].[ScoreAcquisition]"
    Name ="clients.creditFAProfile"
    Expression ="IIf([updateFAProfile].[TOTALSCORE]>=80,\"Qualifies\",IIf([updateFAProfile].[TOTA"
        "LSCORE]>=70,\"Review Required\",\"Does Not Qualify\"))"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="clients.creditType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="clients.fleetSize"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="clients.fleetDaycabs"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="clients.fleetSleepers"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="clients.creditRank"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="clients.creditFAProfile"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="clients.creditFAScore"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScoreCredit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScoreCredit2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScoreFleetSize"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PercentDC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScorePercentDC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.fleetDaycabsMPY"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.fleetSleepersMPY"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScoreDCMiles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.creditMaint"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScoreSLMiles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScoreMAINTENANCE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientState"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientCompanyName"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientType"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.fleetMPY"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.FleetReplacementCycle"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.creditRef"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.creditMoodyRating"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientDBA"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.fleetStraightTrucks"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.fleetStraightTrucksMPY"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.fleetTrailers"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.fleetTrailersMPY"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.creditFinance"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="companyAndDBA"
        dbInteger "ColumnWidth" ="5085"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SalespersonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.[ProspectTarget Sales Volume]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ProspectTargetUnits"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ProspectTOPTEN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ProspectTarget"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.SalespersonREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="view"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScoreLIFECYCLE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScoreACQUISITION"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TOTALSCORE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GRADE"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =163
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =45
        Top =-3
        Right =374
        Bottom =216
        Top =0
        Name ="clients"
        Name =""
    End
End
