Operation =4
Option =0
Where ="(((Units.ObcUnitNumber) Is Null) AND ((Units.UnitStatus)<>\"B\"))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Name ="Units.ObcUnitNumber"
    Expression ="Switch(IsNull([units].[UnitUnitNum]) And [UnitStatus] Like \"*S*\",\"Sold\" & Ri"
        "ght([UnitVin],6),IsNull([units].[UnitUnitNum]) And [UnitStatus]=\"O\",\"OL\" & R"
        "ight([UnitVin],6),IsNull([units].[UnitUnitNum]) And [UnitStatus]=\"T\",\"TRK\" &"
        " Right([UnitVin],6),IsNull([units].[UnitUnitNum]) And [UnitStatus]=\"A\",\"ATL\""
        " & Right([UnitVin],6),Not IsNull([units].[UnitUnitNum]),[UnitUnitNum])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ObcUnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =155
    Top =6
    Right =1627
    Bottom =1005
    Left =-1
    Top =-1
    Right =1440
    Bottom =681
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =116
        Top =237
        Right =456
        Bottom =716
        Top =0
        Name ="Units"
        Name =""
    End
End
