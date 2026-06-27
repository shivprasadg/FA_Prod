Operation =1
Option =1
Where ="(((X.GroupID)=2171) AND ((ATF_ChangeReasons.INDX) Is Null))"
Begin InputTables
    Name ="vw_FAID_RevRentGrouped"
    Alias ="X"
    Name ="ATF_ChangeReasons"
End
Begin OutputColumns
End
Begin Joins
    LeftTable ="X"
    RightTable ="ATF_ChangeReasons"
    Expression ="X.SysUser = ATF_ChangeReasons.SysUser"
    Flag =2
    LeftTable ="X"
    RightTable ="ATF_ChangeReasons"
    Expression ="X.FieldName = ATF_ChangeReasons.FieldName"
    Flag =2
    LeftTable ="X"
    RightTable ="ATF_ChangeReasons"
    Expression ="X.GroupID = ATF_ChangeReasons.GroupID"
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
        dbText "Name" ="ATF_ChangeReasons.INDX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.ActionDate"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.FieldName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.Action"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.SysUser"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.ActionJustDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.OldValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.NewValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRents.ActionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRentGrouped.ActionDate"
        dbInteger "ColumnWidth" ="2130"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.ChangeDate"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRentGrouped.ActionJustDate"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRents.Action"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.RevID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.GroupNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRentGrouped.Action"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRents.SysUser"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRentGrouped.SysUser"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRents.ActionJustDate"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.Action"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRentGrouped.FieldName"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.ReviewedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRentGrouped.OldValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.FieldName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_FAID_RevRentGrouped.NewValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.ATFSchSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.Application"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.tableName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.ActionDesc"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.UnitIDEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.UnitIDStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.FieldLabel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.ATFField"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.FieldGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X.UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.NewValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.GroupNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.FieldGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.OldValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.TableName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.SysUser"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_ChangeReasons.ReviewedBy"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-56
    Top =66
    Right =1379
    Bottom =1052
    Left =-1
    Top =-1
    Right =1417
    Bottom =641
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =507
        Top =0
        Name ="X"
        Name =""
    End
    Begin
        Left =412
        Top =113
        Right =556
        Bottom =401
        Top =0
        Name ="ATF_ChangeReasons"
        Name =""
    End
End
