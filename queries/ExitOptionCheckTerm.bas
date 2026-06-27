Operation =1
Option =0
Where ="(((UnitExitOption.UnitID) Is Null))"
Begin InputTables
    Name ="UnitExitOption"
    Name ="EXITOptions_WorkTable"
End
Begin OutputColumns
    Expression ="UnitExitOption.UnitID"
End
Begin Joins
    LeftTable ="EXITOptions_WorkTable"
    RightTable ="UnitExitOption"
    Expression ="EXITOptions_WorkTable.UnitID = UnitExitOption.UnitID"
    Flag =2
    LeftTable ="EXITOptions_WorkTable"
    RightTable ="UnitExitOption"
    Expression ="EXITOptions_WorkTable.UID2 = UnitExitOption.UIDX"
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
        dbText "Name" ="UnitExitOption.UnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =142
    Top =-10
    Right =1600
    Bottom =568
    Left =-1
    Top =-1
    Right =1426
    Bottom =286
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =552
        Top =33
        Right =696
        Bottom =295
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
    Begin
        Left =338
        Top =156
        Right =482
        Bottom =635
        Top =0
        Name ="EXITOptions_WorkTable"
        Name =""
    End
End
