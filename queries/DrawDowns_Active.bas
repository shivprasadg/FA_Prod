Operation =1
Option =0
Begin InputTables
    Name ="DrawDowns"
    Name ="Units"
    Name ="ChildParts"
    Name ="Banks"
    Name ="MstrLease"
    Name ="DrawDownType"
End
Begin OutputColumns
    Alias ="Bank"
    Expression ="Nz([banks].[BankShortName],\"TBD\")"
    Expression ="DrawDowns.DrawDownId"
    Alias ="Type"
    Expression ="DrawDownType.DrawnDownType"
    Expression ="MstrLease.MLNo"
    Alias ="Request"
    Expression ="DrawDowns.DrawRequestNumber"
    Alias ="Draw"
    Expression ="DrawDowns.DrawNumber"
    Alias ="Units"
    Expression ="Count(Units.Unitid)"
    Alias ="Child"
    Expression ="Count(ChildParts.ChildPartId)"
    Expression ="MstrLease.MLID"
    Expression ="Banks.BankId"
    Expression ="DrawDowns.ClientID"
    Expression ="DrawDowns.DrawNumber"
    Expression ="DrawDowns.DateDrawn"
End
Begin Joins
    LeftTable ="DrawDowns"
    RightTable ="Units"
    Expression ="DrawDowns.DrawDownId = Units.DrawDownId"
    Flag =2
    LeftTable ="DrawDowns"
    RightTable ="ChildParts"
    Expression ="DrawDowns.DrawDownId = ChildParts.DrawDownId"
    Flag =2
    LeftTable ="DrawDowns"
    RightTable ="Banks"
    Expression ="DrawDowns.DrawBankId = Banks.bankID"
    Flag =2
    LeftTable ="DrawDowns"
    RightTable ="MstrLease"
    Expression ="DrawDowns.MLID = MstrLease.MLID"
    Flag =1
    LeftTable ="DrawDowns"
    RightTable ="DrawDownType"
    Expression ="DrawDowns.DrawTypeId = DrawDownType.DrawDownTypeId"
    Flag =1
End
Begin OrderBy
    Expression ="Nz([banks].[BankShortName],\"TBD\")"
    Flag =0
    Expression ="DrawDowns.DrawDownId"
    Flag =0
    Expression ="DrawDowns.DrawRequestNumber"
    Flag =0
End
Begin Groups
    Expression ="Nz([banks].[BankShortName],\"TBD\")"
    GroupLevel =0
    Expression ="DrawDowns.DrawDownId"
    GroupLevel =0
    Expression ="DrawDownType.DrawnDownType"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="DrawDowns.DrawRequestNumber"
    GroupLevel =0
    Expression ="MstrLease.MLID"
    GroupLevel =0
    Expression ="Banks.BankId"
    GroupLevel =0
    Expression ="DrawDowns.ClientID"
    GroupLevel =0
    Expression ="DrawDowns.DrawNumber"
    GroupLevel =0
    Expression ="DrawDowns.DateDrawn"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Request#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Draw#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Child"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DrawNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Draw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Request"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DrawDowns.DateDrawn"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-301
    Top =65
    Right =999
    Bottom =948
    Left =-1
    Top =-1
    Right =1276
    Bottom =481
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =370
        Top =0
        Name ="DrawDowns"
        Name =""
    End
    Begin
        Left =756
        Top =7
        Right =900
        Bottom =151
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =759
        Top =159
        Right =903
        Bottom =303
        Top =0
        Name ="ChildParts"
        Name =""
    End
    Begin
        Left =233
        Top =404
        Right =377
        Bottom =548
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =246
        Top =243
        Right =390
        Bottom =387
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =48
        Top =383
        Right =192
        Bottom =527
        Top =0
        Name ="DrawDownType"
        Name =""
    End
End
