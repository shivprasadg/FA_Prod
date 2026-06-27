Operation =1
Option =0
Having ="(((MstrLease.MLOrig)=\"FA\") AND ((SchGrp.Assigned_Indicator)<>\"F\"))"
Begin InputTables
    Name ="Schedule"
    Name ="SchGrp"
    Name ="MstrLease"
    Name ="Clients"
    Name ="Units"
    Name ="ScheduleAmends"
End
Begin OutputColumns
    Alias ="Lessee"
    Expression ="Clients.ClientShNm"
    Expression ="Schedule.SchNo"
    Alias ="EarliestDelvDate"
    Expression ="Min(Units.UnitActualDelDate)"
    Alias ="LatestDelvDate"
    Expression ="Max(Units.UnitActualDelDate)"
    Alias ="SchOutBack"
    Expression ="IIf(IsDate([Schout]),\"Out  \" & [SchOut] & Chr(13) & Chr(10),Null) & IIf(IsDate"
        "([SchIn]),\"Back \" & [SchIn],\"Pending\")"
    Alias ="FinalEconRecv"
    Expression ="IIf(IsDate([SchFEcoRec]),[SchFEcoRec],\"Pending\")"
    Alias ="RALOutBack"
    Expression ="IIf(IsDate([SchRALO]),\"Out \" & [SchRALO] & Chr(13) & Chr(10),Null) & IIf(IsDat"
        "e([SchRALB]),\"Back  \" & [SchRALB],\"Pending\")"
    Alias ="OPSPkg"
    Expression ="IIf(IsDate([SchOpsPkg]),[SchOpsPkg],\"Pending\")"
    Alias ="PermPkgOut"
    Expression ="IIf(IsDate([SchFPkgO]),\"Out\",Null)"
    Alias ="FATitleTrust"
    Expression ="IIf([SchTitleToTrust]=True,\"Yes\",\"No\")"
    Alias ="IntermFundBank"
    Expression ="Max(Nz([UnitIFBank],\"n/a\"))"
    Alias ="IntermFund"
    Expression ="Max(Nz([UnitPayOff],\"n/a\"))"
    Alias ="BLCD"
    Expression ="\"Groups:\" & Chr(13) & Chr(10) & BLCDInSchedule([SchId])"
    Alias ="AssignmentDate"
    Expression ="\"Groups:\" & Chr(13) & Chr(10) & AssignmentDateInSchedule([SchId])"
    Alias ="Assignee"
    Expression ="SchGrp.Assignee_Bank"
    Alias ="FundDate"
    Expression ="\"Groups:\" & Chr(13) & Chr(10) & FundDateInSchedule([SchId])"
    Expression ="Schedule.SchID"
    Alias ="GroupID"
    Expression ="Min(SchGrp.SGrpID)"
    Expression ="Clients.clientID"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="Schedule"
    Expression ="SchGrp.SGrpSchID = Schedule.SchID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="MstrLease"
    Expression ="Schedule.SchMLRef = MstrLease.MLID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="ScheduleAmends"
    Expression ="Schedule.SchID = ScheduleAmends.ScheduleID"
    Flag =2
End
Begin Groups
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="IIf(IsDate([Schout]),\"Out  \" & [SchOut] & Chr(13) & Chr(10),Null) & IIf(IsDate"
        "([SchIn]),\"Back \" & [SchIn],\"Pending\")"
    GroupLevel =0
    Expression ="IIf(IsDate([SchFEcoRec]),[SchFEcoRec],\"Pending\")"
    GroupLevel =0
    Expression ="IIf(IsDate([SchRALO]),\"Out \" & [SchRALO] & Chr(13) & Chr(10),Null) & IIf(IsDat"
        "e([SchRALB]),\"Back  \" & [SchRALB],\"Pending\")"
    GroupLevel =0
    Expression ="IIf(IsDate([SchOpsPkg]),[SchOpsPkg],\"Pending\")"
    GroupLevel =0
    Expression ="IIf(IsDate([SchFPkgO]),\"Out\",Null)"
    GroupLevel =0
    Expression ="IIf([SchTitleToTrust]=True,\"Yes\",\"No\")"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="Clients.clientID"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
    GroupLevel =0
    Expression ="SchGrp.Assigned_Indicator"
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
dbInteger "RowHeight" ="1245"
Begin
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="FATitleTrust"
        dbText "Format" ="Yes/No"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="AssignmentDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="FundDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="IntermFund"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="IntermFundBank"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="EarliestDelvDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="LatestDelvDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Lessee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="BLCD"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="RALOutBack"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="OPSPkg"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="FinalEconRecv"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="SchOutBack"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="PermPkgOut"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =64
    Top =55
    Right =1380
    Bottom =980
    Left =-1
    Top =-1
    Right =1298
    Bottom =486
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =212
        Top =71
        Right =356
        Bottom =215
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =398
        Top =135
        Right =542
        Bottom =640
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =23
        Top =18
        Right =167
        Bottom =162
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =48
        Top =244
        Right =192
        Bottom =388
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =584
        Top =129
        Right =728
        Bottom =608
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =806
        Top =52
        Right =950
        Bottom =288
        Top =0
        Name ="ScheduleAmends"
        Name =""
    End
End
