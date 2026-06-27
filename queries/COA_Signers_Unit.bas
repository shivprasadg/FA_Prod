Operation =1
Option =0
Having ="(((Units.CofASigner) Is Not Null And (Units.CofASigner)>0))"
Begin InputTables
    Name ="Units"
    Name ="SchGrp"
End
Begin OutputColumns
    Alias ="SignerType"
    Expression ="\"COA_Signer\""
    Alias ="SchID"
    Expression ="SchGrp.SGrpSchID"
    Expression ="SchGrp.SGrpID"
    Alias ="CofASignerID"
    Expression ="Units.CofASigner"
    Expression ="SchGrp.SGrpNO"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="SchGrp"
    Expression ="Units.GroupREF = SchGrp.SGrpID"
    Flag =1
End
Begin Groups
    Expression ="SchGrp.SGrpSchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="Units.CofASigner"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
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
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SignerType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofASignerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =197
    Top =45
    Right =1323
    Bottom =908
    Left =-1
    Top =-1
    Right =1108
    Bottom =569
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =454
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
