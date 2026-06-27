Operation =1
Option =8
Having ="(((MstrLease.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="COA_Signers_Unit"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Clients.ClientShNm"
    Expression ="MstrLease.MLNo"
    Expression ="MstrLease.MLOrig"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Expression ="COA_Signers_Unit.SGrpID"
    Expression ="COA_Signers_Unit.SGrpNO"
    Expression ="MstrLease.MLSignerID"
    Expression ="MstrLease.MLSigner2ID"
    Expression ="MstrLease.MLSigner3ID"
    Expression ="MstrLease.MLGuarantor"
    Expression ="MstrLease.MLGuarSignerID"
    Alias ="SignerGroup1"
    Expression ="IIf(Nz([MLSignerID],0)+Nz([MLSigner2ID],0)+Nz([MLSigner3ID],0)=0,Null,\"Master L"
        "ease\")"
    Alias ="SignerGroup2"
    Expression ="IIf(IsNull([MLGuarantor]),Null,\"Guarantor\")"
    Alias ="SignerGroup3"
    Expression ="IIf(Nz([SchSignerId],0)+Nz([SchSigner2ID],0)+Nz([SchSigner3ID],0)=0,Null,\"Sched"
        "ule\")"
    Alias ="SignerGroup4"
    Expression ="IIf(IsNull([CofASignerID]),Null,\"COA\")"
    Expression ="Schedule.SchSignerID"
    Expression ="Schedule.SchSigner2ID"
    Expression ="Schedule.SchSigner3ID"
    Expression ="COA_Signers_Unit.CofASignerID"
    Expression ="MstrLease.MLID"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="COA_Signers_Unit"
    Expression ="Schedule.SchID = COA_Signers_Unit.SchID"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="COA_Signers_Unit.SGrpID"
    Flag =0
End
Begin Groups
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="COA_Signers_Unit.SGrpID"
    GroupLevel =0
    Expression ="COA_Signers_Unit.SGrpNO"
    GroupLevel =0
    Expression ="MstrLease.MLSignerID"
    GroupLevel =0
    Expression ="MstrLease.MLSigner2ID"
    GroupLevel =0
    Expression ="MstrLease.MLSigner3ID"
    GroupLevel =0
    Expression ="MstrLease.MLGuarantor"
    GroupLevel =0
    Expression ="MstrLease.MLGuarSignerID"
    GroupLevel =0
    Expression ="IIf(Nz([MLSignerID],0)+Nz([MLSigner2ID],0)+Nz([MLSigner3ID],0)=0,Null,\"Master L"
        "ease\")"
    GroupLevel =0
    Expression ="IIf(IsNull([MLGuarantor]),Null,\"Guarantor\")"
    GroupLevel =0
    Expression ="IIf(Nz([SchSignerId],0)+Nz([SchSigner2ID],0)+Nz([SchSigner3ID],0)=0,Null,\"Sched"
        "ule\")"
    GroupLevel =0
    Expression ="IIf(IsNull([CofASignerID]),Null,\"COA\")"
    GroupLevel =0
    Expression ="Schedule.SchSignerID"
    GroupLevel =0
    Expression ="Schedule.SchSigner2ID"
    GroupLevel =0
    Expression ="Schedule.SchSigner3ID"
    GroupLevel =0
    Expression ="COA_Signers_Unit.CofASignerID"
    GroupLevel =0
    Expression ="MstrLease.MLID"
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
dbMemo "Filter" ="([ATF_FAID_MLSigners].[SchID]=251)"
Begin
    Begin
        dbText "Name" ="Clients.clientID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnWidth" ="2475"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLGuarantor"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLSignerID"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLSigner2ID"
        dbInteger "ColumnWidth" ="1575"
        dbInteger "ColumnOrder" ="16"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLSigner3ID"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLGuarSignerID"
        dbInteger "ColumnWidth" ="1905"
        dbInteger "ColumnOrder" ="14"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchSignerID"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchSigner2ID"
        dbInteger "ColumnWidth" ="1575"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SignerGroup1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SignerGroup3"
        dbInteger "ColumnWidth" ="1950"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchSigner3ID"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SignerGroup2"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="SignerGroup4"
        dbInteger "ColumnWidth" ="1950"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COA_Signers_Unit.CofASignerID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="COA_Signers_Unit.SGrpID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="COA_Signers_Unit.SGrpNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =27
    Top =53
    Right =1827
    Bottom =261
    Left =-1
    Top =-1
    Right =1782
    Bottom =133
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =414
        Bottom =421
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =447
        Top =91
        Right =591
        Bottom =235
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =639
        Top =12
        Right =783
        Bottom =156
        Top =0
        Name ="COA_Signers_Unit"
        Name =""
    End
End
