Operation =1
Option =0
Begin InputTables
    Name ="SchGrp"
    Name ="vw_SixKeys"
    Name ="Banks"
End
Begin OutputColumns
    Expression ="Banks.BankId"
    Alias ="Assignee"
    Expression ="IIf([mlorig]<>\"FA\" And IsNull([banks].[BankShortName]),\"Not FA\",[banks].[Ban"
        "kShortName])"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.AssetID)"
    Expression ="vw_SixKeys.MLID"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="vw_SixKeys"
    Expression ="SchGrp.SGrpID = vw_SixKeys.SGrpID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Banks"
    Expression ="SchGrp.Assignee_Bank = Banks.BankShortName"
    Flag =2
End
Begin Groups
    Expression ="Banks.BankId"
    GroupLevel =0
    Expression ="IIf([mlorig]<>\"FA\" And IsNull([banks].[BankShortName]),\"Not FA\",[banks].[Ban"
        "kShortName])"
    GroupLevel =0
    Expression ="vw_SixKeys.MLID"
    GroupLevel =0
    Expression ="vw_SixKeys.SchID"
    GroupLevel =0
    Expression ="vw_SixKeys.SGrpID"
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
dbMemo "Filter" ="([AssigneeBanks].[Assignee] Not In (\"Not FA\"))"
Begin
    Begin
        dbText "Name" ="Units"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankId"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1230
    Bottom =1021
    Left =-1
    Top =-1
    Right =1206
    Bottom =647
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =238
        Top =70
        Right =384
        Bottom =462
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =405
        Top =153
        Right =549
        Bottom =297
        Top =0
        Name ="Banks"
        Name =""
    End
End
