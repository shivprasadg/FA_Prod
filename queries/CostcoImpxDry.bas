Operation =3
Name ="TrailerMilesHours"
Option =0
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Costco_DRY"
End
Begin OutputColumns
    Name ="Vin"
    Expression ="vw_SixKeys.Vin"
    Name ="UnitNumber"
    Expression ="vw_SixKeys.Unitnum"
    Name ="LogActivityDate"
    Expression ="Costco_DRY.LogActivityDate"
    Name ="miles"
    Expression ="Costco_DRY.Miles"
    Name ="hours"
    Expression ="Costco_DRY.Hours"
    Alias ="Expr1"
    Name ="ReadDate"
    Expression ="DateAdd(\"d\",33,[LogActivityDate])"
    Name ="UnitID"
    Expression ="vw_SixKeys.UnitID"
    Alias ="Expr2"
    Name ="OBCDataStatus"
    Expression ="1"
    Alias ="Expr3"
    Name ="OBCDataDesc"
    Expression ="\"Reporting\""
    Alias ="Expr4"
    Name ="MaintLeaseType"
    Expression ="\"FSL\""
    Alias ="Expr6"
    Name ="CPM"
    Expression ="0"
    Alias ="Expr5"
    Name ="CPH"
    Expression ="0"
    Name ="milesltd"
    Expression ="Costco_DRY.milesltd"
    Name ="hoursltd"
    Expression ="Costco_DRY.hoursltd"
    Alias ="Expr7"
    Name ="ProviderID"
    Expression ="103"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Costco_DRY"
    Expression ="vw_SixKeys.Unitnum = Costco_DRY.UnitNum"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Costco_DRY.Vin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_DRY.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_DRY.mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_DRY.hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Vin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_DRY.hoursltd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_DRY.milesltd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_DRY.Miles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_Wet.hoursltd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Unitnum"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =96
    Top =0
    Right =1633
    Bottom =965
    Left =-1
    Top =-1
    Right =1459
    Bottom =641
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =167
        Top =65
        Right =311
        Bottom =438
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =359
        Top =12
        Right =503
        Bottom =156
        Top =0
        Name ="Costco_DRY"
        Name =""
    End
End
