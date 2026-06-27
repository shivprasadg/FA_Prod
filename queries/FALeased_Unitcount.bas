Operation =1
Option =0
Having ="(((VW_Clients_ML_FAID_SCH_Groups_count.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="VW_Clients_ML_FAID_SCH_Groups_count"
End
Begin OutputColumns
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.ClientGroupId"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.clientID"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.ClientShortName"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.CompanyName"
    Alias ="UnitCount"
    Expression ="Sum(VW_Clients_ML_FAID_SCH_Groups_count.Units)"
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.MLOrig"
End
Begin Groups
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.ClientGroupId"
    GroupLevel =0
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.clientID"
    GroupLevel =0
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.ClientShortName"
    GroupLevel =0
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.CompanyName"
    GroupLevel =0
    Expression ="VW_Clients_ML_FAID_SCH_Groups_count.MLOrig"
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
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.CompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VW_Clients_ML_FAID_SCH_Groups_count.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =132
    Top =199
    Right =1319
    Bottom =974
    Left =-1
    Top =-1
    Right =1155
    Bottom =457
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =102
        Top =104
        Right =481
        Bottom =421
        Top =0
        Name ="VW_Clients_ML_FAID_SCH_Groups_count"
        Name =""
    End
End
