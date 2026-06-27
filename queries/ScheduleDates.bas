Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
    Name ="TMTSignedDates"
End
Begin OutputColumns
    Alias ="SchID"
    Expression ="SchGrp.SGrpSchID"
    Alias ="GBLCD"
    Expression ="[SGrpNO] & \" - \" & Format([BLCD],\"mm/dd/yy\")"
    Alias ="GFundDate"
    Expression ="[SGrpNO] & \" - \" & Format([FundDate],\"mm/dd/yy\")"
    Alias ="GAssignmentDate"
    Expression ="[SGrpNO] & \" - \" & Format([AssignmentDate],\"mm/dd/yy\")"
    Alias ="GTerm"
    Expression ="[SGrpNO] & \" - \" & [LeaseTermPrimary]"
    Alias ="GMiles"
    Expression ="[SGrpNO] & \" - \" & [SGprAllow1] & \" | \" & [SGprAllow2]"
    Alias ="GRALOut"
    Expression ="[SGrpNO] & \" - \" & IIf(IsDate([sgrpralo]),Format([SGrpRALo],\"mm/dd/yy\"),\"\""
        ")"
    Alias ="GRALBack"
    Expression ="[SGrpNO] & \" - \" & IIf(IsDate([sgrpralb]),Format([SGrpRALB],\"mm/dd/yy\"),\"\""
        ")"
    Alias ="Assignee"
    Expression ="SchGrp.Assignee_Bank"
    Expression ="TMTSignedDates.TMTSigned"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="TMTSignedDates"
    Expression ="SchGrp.SGrpID = TMTSignedDates.GroupID"
    Flag =2
End
Begin Groups
    Expression ="SchGrp.SGrpSchID"
    GroupLevel =0
    Expression ="[SGrpNO] & \" - \" & Format([BLCD],\"mm/dd/yy\")"
    GroupLevel =0
    Expression ="[SGrpNO] & \" - \" & Format([FundDate],\"mm/dd/yy\")"
    GroupLevel =0
    Expression ="[SGrpNO] & \" - \" & Format([AssignmentDate],\"mm/dd/yy\")"
    GroupLevel =0
    Expression ="[SGrpNO] & \" - \" & [LeaseTermPrimary]"
    GroupLevel =0
    Expression ="[SGrpNO] & \" - \" & [SGprAllow1] & \" | \" & [SGprAllow2]"
    GroupLevel =0
    Expression ="[SGrpNO] & \" - \" & IIf(IsDate([sgrpralo]),Format([SGrpRALo],\"mm/dd/yy\"),\"\""
        ")"
    GroupLevel =0
    Expression ="[SGrpNO] & \" - \" & IIf(IsDate([sgrpralb]),Format([SGrpRALB],\"mm/dd/yy\"),\"\""
        ")"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
    GroupLevel =0
    Expression ="TMTSignedDates.TMTSigned"
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
dbInteger "RowHeight" ="855"
Begin
    Begin
        dbText "Name" ="GFundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GAssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GMiles"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GRALOut"
        dbInteger "ColumnWidth" ="3225"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GRALBack"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TMTSignedDates.TMTSigned"
        dbInteger "ColumnWidth" ="1440"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =45
    Right =1465
    Bottom =868
    Left =-1
    Top =-1
    Right =1397
    Bottom =223
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =14
        Top =33
        Right =268
        Bottom =458
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =411
        Top =60
        Right =555
        Bottom =204
        Top =0
        Name ="TMTSignedDates"
        Name =""
    End
End
