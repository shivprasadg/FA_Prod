Operation =1
Option =0
Begin InputTables
    Name ="vw_UMP_Mileage"
End
Begin OutputColumns
    Expression ="vw_UMP_Mileage.*"
End
Begin OrderBy
    Expression ="vw_UMP_Mileage.LogActivityDate"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vw_UMP_Mileage.UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.Miles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.Fuel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.FuelMoving"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.FuelIdle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.MPGMoving"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.OdometerBegin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.OdometerEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.MPGOverAll"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.MPGLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.OBCDataStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.OBCDataStatusDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.FuelMovingLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.MilesLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.MPGGoalDriving"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.FuelLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.MovingTimePercentage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.DataProviderID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.HoursLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UMP_Mileage.MPGGoalOverall"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =13
    Top =66
    Right =1601
    Bottom =973
    Left =-1
    Top =-1
    Right =1556
    Bottom =454
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =332
        Bottom =298
        Top =0
        Name ="vw_UMP_Mileage"
        Name =""
    End
End
