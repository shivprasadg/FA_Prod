Operation =1
Option =0
Where ="(((Units.ObcUnitNumber) In (SELECT [ObcUnitNumber] FROM [Units] As Tmp GROUP BY "
    "[ObcUnitNumber],[ClientGroupId] HAVING Count(*)>1  And [ClientGroupId] = [Units]"
    ".[ClientGroupId])) AND ((Units.ClientGroupId)=697))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.ObcUnitNumber"
    Expression ="Units.ClientGroupId"
    Expression ="Units.UnitID"
    Expression ="Units.UnitYr"
    Expression ="Units.ClientID"
    Expression ="Units.UnitMake"
    Expression ="Units.UnitModel"
    Expression ="Units.UnitType"
End
Begin OrderBy
    Expression ="Units.ObcUnitNumber"
    Flag =0
    Expression ="Units.ClientGroupId"
    Flag =0
    Expression ="Units.UnitYr"
    Flag =1
    Expression ="Units.ClientID"
    Flag =0
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
        dbText "Name" ="Units.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[ClientID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[ObcUnitNumber]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[ClientGroupId]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.[UnitID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ObcUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =149
    Top =132
    Right =1285
    Bottom =1014
    Left =-1
    Top =-1
    Right =1104
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =221
        Top =0
        Name ="Units"
        Name =""
    End
End
