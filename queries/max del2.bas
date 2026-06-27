Operation =1
Option =0
Having ="(((MstrLease.MLOrig)=\"fa\") AND ((MstrLease.MLNo) Not Like \"demo*\")) OR (((Ms"
    "trLease.MLOrig)=\"fa\") AND ((MstrLease.MLNo) Not Like \"demo*\"))"
Begin InputTables
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="SchGrp.SGrpSchID"
    Expression ="SchGrp.SGrpID"
    Alias ="ProjSyn"
    Expression ="Max(DateAdd(\"m\",1,IIf(IsNull([unitinserv]),IIf(IsNull([unitacceptdate]),IIf(Is"
        "Null([unitestdeldate]),[UNITESTDELDATE],[unitestdeldate]),[unitacceptdate]),[uni"
        "tinserv]))-1)"
    Expression ="MstrLease.MLOrig"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchNo"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.LoanType"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="SchGrp.FundDate"
    Expression ="SchGrp.LoanBank"
    Alias ="AccptUnit"
    Expression ="Sum(IIf(([unitstatus]=\"a\" And [unitdocaccpt]=No) Or [unitinservind]=\"a\",1,0)"
        ")"
    Alias ="DocsAccpt"
    Expression ="Sum(IIf(IsNull([UnitAcceptDate]),0,IIf([unitdocaccpt]=Yes And [unitinservind]=\""
        "p\",1,0)))"
    Alias ="MaxOfUnitBuildDate"
    Expression ="Max(Units.UnitBuildDate)"
End
Begin Joins
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchNo"
    Flag =0
    Expression ="SchGrp.SGrpNO"
    Flag =0
End
Begin Groups
    Expression ="SchGrp.SGrpSchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="MstrLease.MLOrig"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="SchGrp.LoanType"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
    GroupLevel =0
    Expression ="SchGrp.FundDate"
    GroupLevel =0
    Expression ="SchGrp.LoanBank"
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
        dbText "Name" ="ProjSyn"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpfunddt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpLoanTpe"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpAssignee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="SchGrp.sgrploanbank"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="AccptUnit"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="DocsAccpt"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfUnitBuildDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3645"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SGrpFunddt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee_Bank"
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
    Bottom =312
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =78
        Top =2
        Right =222
        Bottom =146
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =274
        Top =-5
        Right =418
        Bottom =139
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =402
        Top =-5
        Right =546
        Bottom =139
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =590
        Top =-9
        Right =744
        Bottom =146
        Top =0
        Name ="Units"
        Name =""
    End
End
