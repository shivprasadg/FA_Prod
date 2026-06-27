Operation =1
Option =0
Begin InputTables
    Name ="SchedulesSubform_Pre"
    Name ="ScheduleExchangedBy"
End
Begin OutputColumns
    Expression ="SchedulesSubform_Pre.*"
    Expression ="ScheduleExchangedBy.NewSchGrp"
    Expression ="ScheduleExchangedBy.ReplacedByScheduleID"
    Expression ="ScheduleExchangedBy.ReplacedByGroupID"
End
Begin Joins
    LeftTable ="SchedulesSubform_Pre"
    RightTable ="ScheduleExchangedBy"
    Expression ="SchedulesSubform_Pre.SchID = ScheduleExchangedBy.SchID"
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
        dbText "Name" ="ScheduleExchangedBy.ReplacedByGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchIn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchSigner2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchSigner3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchAssDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchTitleToTrust"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SCHINSAMT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Edit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Bassign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchTitleAssignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SCHINREC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchAlias"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.Schsigned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.LienHolder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.LienHolderNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.MstrLease.MLLessor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchAH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SCHNo1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchMLRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchSigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchTitltoFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SchTitletoClient"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.MstrLease.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.Schedule.SCHINSEXP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleExchangedBy.ReplacedByScheduleID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.ECost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleExchangedBy.NewSchGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchedulesSubform_Pre.MaxOfUnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1527
    Bottom =991
    Left =-1
    Top =-1
    Right =1495
    Bottom =650
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =100
        Top =106
        Right =244
        Bottom =502
        Top =0
        Name ="SchedulesSubform_Pre"
        Name =""
    End
    Begin
        Left =360
        Top =194
        Right =504
        Bottom =338
        Top =0
        Name ="ScheduleExchangedBy"
        Name =""
    End
End
