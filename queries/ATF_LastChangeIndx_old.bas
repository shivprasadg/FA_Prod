Operation =1
Option =0
Having ="(((ATF_ChangeReasons.FieldName) Is Not Null))"
Begin InputTables
    Name ="ATF_ChangeReasons"
End
Begin OutputColumns
    Alias ="INDX"
    Expression ="Max(ATF_ChangeReasons.Indx)"
    Expression ="ATF_ChangeReasons.FieldName"
    Expression ="ATF_ChangeReasons.GroupID"
    Expression ="ATF_ChangeReasons.SysUser"
End
Begin OrderBy
    Expression ="ATF_ChangeReasons.FieldName"
    Flag =0
End
Begin Groups
    Expression ="ATF_ChangeReasons.FieldName"
    GroupLevel =0
    Expression ="ATF_ChangeReasons.GroupID"
    GroupLevel =0
    Expression ="ATF_ChangeReasons.SysUser"
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
        dbText "Name" ="INDX"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =210
    Top =59
    Right =1332
    Bottom =983
    Left =-1
    Top =-1
    Right =1104
    Bottom =596
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =119
        Top =78
        Right =263
        Bottom =396
        Top =0
        Name ="ATF_ChangeReasons"
        Name =""
    End
End
