Operation =3
Name ="ATFScheduleSpec"
Option =0
Where ="(((ATF_SF_ScheduleProfile.FAID)=[FAIDx]))"
Begin InputTables
    Name ="ATF_SF_ScheduleProfile"
End
Begin OutputColumns
    Alias ="ATFSchName"
    Name ="ATFSchName"
    Expression ="ATF_SF_ScheduleProfile.SchName"
    Alias ="SchFAID"
    Name ="SchFAID"
    Expression ="ATF_SF_ScheduleProfile.FAID"
    Alias ="ATFID"
    Name ="ATFID"
    Expression ="[ATFX]"
    Alias ="MLID"
    Name ="MLID"
    Expression ="[MLIDx]"
    Alias ="FSLType"
    Name ="FSLType"
    Expression ="\"Yes\""
    Alias ="ATFSchSpecDesc"
    Name ="ATFSchSpecDesc"
    Expression ="ATF_SF_ScheduleProfile.SchName"
    Alias ="SchDescription"
    Name ="SchDescription"
    Expression ="\"(\" & DSum(\"UnitsInGroup\",\"ATF_SF_GroupProfile\",\"SchName='\" & [SchName] "
        "& \"' AND GrpFAID='\" & [FAIDx] & \"'\") & \") \" & SplitString([SchShortDesc],\""
        ") \",1)"
    Alias ="NumGroupsInSch"
    Name ="NumGroupsInSch"
    Expression ="DLookUp(\"GroupsIn\",\"ATF_SF_GroupProfile\",\"SchName='\" & [SchName] & \"' AND"
        " GrpFAID='15-8168 '\")"
    Alias ="SchSpecInserted"
    Name ="SchSpecInserted"
    Expression ="Now()"
End
Begin Parameters
    Name ="ATFX"
    Flag =4
    Name ="MLIDx"
    Flag =4
    Name ="FAIDx"
    Flag =10
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.LOExec"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.ShortGUID"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.ClientGroupId"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.TaxID"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.clientID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.FAID"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.BDEX"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.LesseeID"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.Lessee"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsInEachSch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.SchName"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.SchShortDesc"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_SF_ScheduleProfile.DealUnitQty"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumGroupsInSch"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="FSLType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="MLID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFSchSpecDesc"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchDescription"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchFAID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFSchName"
        dbInteger "ColumnWidth" ="2550"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchSpecInserted"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =105
    Top =33
    Right =1582
    Bottom =544
    Left =-1
    Top =-1
    Right =1459
    Bottom =198
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =64
        Top =27
        Right =208
        Bottom =171
        Top =0
        Name ="ATF_SF_ScheduleProfile"
        Name =""
    End
End
