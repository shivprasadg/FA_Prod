Operation =1
Option =0
Having ="(((CDate(Format(Nz([ATF_ChangeReasons].[ChangeDate],\"1/1/1900\"),\"mm/dd/yyyy\""
    ")))=Date()))"
Begin InputTables
    Name ="ATF_ChangeReasons"
    Name ="ATF_LastChangeIndx"
End
Begin OutputColumns
    Expression ="ATF_ChangeReasons.INDX"
    Expression ="ATF_ChangeReasons.FieldName"
    Expression ="ATF_ChangeReasons.GroupID"
    Expression ="ATF_ChangeReasons.SysUser"
    Alias ="ChangeDate"
    Expression ="CDate(Format(Nz([ATF_ChangeReasons].[ChangeDate],\"1/1/1900\"),\"mm/dd/yyyy\"))"
    Alias ="ChangeDateTime"
    Expression ="ATF_ChangeReasons.ChangeDate"
    Expression ="ATF_ChangeReasons.Reason"
End
Begin Joins
    LeftTable ="ATF_ChangeReasons"
    RightTable ="ATF_LastChangeIndx"
    Expression ="ATF_ChangeReasons.INDX = ATF_LastChangeIndx.INDX"
    Flag =1
End
Begin OrderBy
    Expression ="ATF_ChangeReasons.FieldName"
    Flag =0
End
Begin Groups
    Expression ="ATF_ChangeReasons.INDX"
    GroupLevel =0
    Expression ="ATF_ChangeReasons.FieldName"
    GroupLevel =0
    Expression ="ATF_ChangeReasons.GroupID"
    GroupLevel =0
    Expression ="ATF_ChangeReasons.SysUser"
    GroupLevel =0
    Expression ="ATF_ChangeReasons.ChangeDate"
    GroupLevel =0
    Expression ="ATF_ChangeReasons.Reason"
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
        dbText "Name" ="ATF_ChangeReasons.FieldName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.SysUser"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChangeDate"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChangeDateTime"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.INDX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.Reason"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =106
    Top =172
    Right =1228
    Bottom =1096
    Left =-1
    Top =-1
    Right =1104
    Bottom =409
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =239
        Top =62
        Right =383
        Bottom =560
        Top =0
        Name ="ATF_ChangeReasons"
        Name =""
    End
    Begin
        Left =38
        Top =91
        Right =182
        Bottom =235
        Top =0
        Name ="ATF_LastChangeIndx"
        Name =""
    End
End
