Operation =1
Option =0
Having ="(((MstrLease.MLID)=[Forms]![ClientsView]![LeaseSum].[Form]![MLID]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchEquipCostq"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="MstrLease.MLOrig"
    Expression ="MstrLease.MLID"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchFAID"
    Expression ="Schedule.SCHNo1"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchFSL"
    Expression ="Schedule.SchMLRef"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.SchAlias"
    Expression ="Schedule.SchOut"
    Expression ="Schedule.SchIn"
    Expression ="Schedule.Schsigned"
    Expression ="Schedule.SchSigner"
    Expression ="Schedule.SchSigner2"
    Expression ="Schedule.SchSigner3"
    Expression ="Schedule.SchAssDt"
    Expression ="Schedule.SchTitleAssignee"
    Expression ="Schedule.SchTitleToTrust"
    Expression ="Schedule.SchShortDesc"
    Expression ="Schedule.SchFAID"
    Expression ="Schedule.SCHINSAMT"
    Expression ="Schedule.SCHINREC"
    Expression ="Schedule.SCHINSEXP"
    Expression ="Schedule.SchTitltoFA"
    Expression ="Schedule.SchTitletoClient"
    Expression ="Schedule.LienHolder"
    Expression ="Schedule.LienHolderNm"
    Expression ="MstrLease.MLLessor"
    Expression ="MstrLease.Client"
    Expression ="Clients.ClientShNm"
    Alias ="Edit"
    Expression ="\"Edit\""
    Expression ="Schedule.SchAH"
    Alias ="Bassign"
    Expression ="IIf(IsNull([Assignee_Bank\015\012\015\012\015\012\015\012]),Null,[Assignee_Bank\015"
        "\012\015\012\015\012\015\012])"
    Alias ="ECost"
    Expression ="IIf(IsNull([SumOfOrigEquipCost]),Null,[SumOfOrigEquipCost])"
    Alias ="MaxOfUnitEstDelDate"
    Expression ="Max(Units.UnitEstDelDate)"
    Alias ="PLB"
    Expression ="IIf(IsNull([SgrpPlb]),No,[SgrpPlb])"
    Expression ="SchGrp.Assigned_Indicator"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchEquipCostq"
    Expression ="Schedule.SchID = SchEquipCostq.SchID"
    Flag =2
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =2
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SCHNo1"
    Flag =1
    Expression ="Schedule.SchID"
    Flag =1
    Expression ="Schedule.SchNo"
    Flag =1
End
Begin Groups
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
    GroupLevel =0
    Expression ="MstrLease.MLID"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SCHNo1"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Schedule.SchFSL"
    GroupLevel =0
    Expression ="Schedule.SchMLRef"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="Schedule.SchAlias"
    GroupLevel =0
    Expression ="Schedule.SchOut"
    GroupLevel =0
    Expression ="Schedule.SchIn"
    GroupLevel =0
    Expression ="Schedule.Schsigned"
    GroupLevel =0
    Expression ="Schedule.SchSigner"
    GroupLevel =0
    Expression ="Schedule.SchSigner2"
    GroupLevel =0
    Expression ="Schedule.SchSigner3"
    GroupLevel =0
    Expression ="Schedule.SchAssDt"
    GroupLevel =0
    Expression ="Schedule.SchTitleAssignee"
    GroupLevel =0
    Expression ="Schedule.SchTitleToTrust"
    GroupLevel =0
    Expression ="Schedule.SchShortDesc"
    GroupLevel =0
    Expression ="Schedule.SchFAID"
    GroupLevel =0
    Expression ="Schedule.SCHINSAMT"
    GroupLevel =0
    Expression ="Schedule.SCHINREC"
    GroupLevel =0
    Expression ="Schedule.SCHINSEXP"
    GroupLevel =0
    Expression ="Schedule.SchTitltoFA"
    GroupLevel =0
    Expression ="Schedule.SchTitletoClient"
    GroupLevel =0
    Expression ="Schedule.LienHolder"
    GroupLevel =0
    Expression ="Schedule.LienHolderNm"
    GroupLevel =0
    Expression ="MstrLease.MLLessor"
    GroupLevel =0
    Expression ="MstrLease.Client"
    GroupLevel =0
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="\"Edit\""
    GroupLevel =0
    Expression ="Schedule.SchAH"
    GroupLevel =0
    Expression ="IIf(IsNull([Assignee_Bank\015\012\015\012\015\012\015\012]),Null,[Assignee_Bank\015"
        "\012\015\012\015\012\015\012])"
    GroupLevel =0
    Expression ="IIf(IsNull([SumOfOrigEquipCost]),Null,[SumOfOrigEquipCost])"
    GroupLevel =0
    Expression ="IIf(IsNull([SgrpPlb]),No,[SgrpPlb])"
    GroupLevel =0
    Expression ="SchGrp.Assigned_Indicator"
    GroupLevel =0
    Expression ="Schedule.SchFAID"
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
        dbText "Name" ="Schedule.SchMLRef"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="MstrLease.MLLessor"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="Schedule.LienHolderNm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Schedule.SchAlias"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Schedule.SchOut"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Schedule.SchIn"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Schedule.Schsigned"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Schedule.SchSigner"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5325"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Schedule.SchSigner2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Schedule.SchSigner3"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Schedule.SchAssDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Schedule.SchTitleAssignee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="Schedule.SchTitleToTrust"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3450"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="Schedule.SchFAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Schedule.SCHINSAMT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="Schedule.SCHINREC"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Schedule.SCHINSEXP"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Schedule.SchTitltoFA"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="Schedule.SchTitletoClient"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="Schedule.LienHolder"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="MstrLease.Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Edit"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="Schedule.SchAH"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="Expr1004"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Bassign"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="ECost"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="MaxOfUnitEstDelDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="PLB"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="38"
    End
    Begin
        dbText "Name" ="Schedule.SCHNo1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1155"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpAssInd"
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
    Bottom =282
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =8
        Right =192
        Bottom =152
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =197
        Top =9
        Right =341
        Bottom =153
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =347
        Top =7
        Right =490
        Bottom =176
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =520
        Top =50
        Right =622
        Bottom =194
        Top =0
        Name ="SchEquipCostq"
        Name =""
    End
    Begin
        Left =632
        Top =25
        Right =776
        Bottom =169
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =806
        Top =95
        Right =950
        Bottom =239
        Top =0
        Name ="Units"
        Name =""
    End
End
