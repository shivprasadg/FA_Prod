Operation =3
Name ="TrailerMilesHours"
Option =0
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Costco_Wet"
End
Begin OutputColumns
    Name ="Vin"
    Expression ="vw_SixKeys.Vin"
    Name ="UnitNumber"
    Expression ="vw_SixKeys.Unitnum"
    Name ="LogActivityDate"
    Expression ="Costco_Wet.LogActivityDate"
    Name ="miles"
    Expression ="Costco_Wet.Miles"
    Name ="hours"
    Expression ="Costco_Wet.Hours"
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
    Expression ="Costco_Wet.milesltd"
    Name ="hoursltd"
    Expression ="Costco_Wet.hoursltd"
    Alias ="Expr7"
    Name ="ProviderID"
    Expression ="103"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Costco_Wet"
    Expression ="vw_SixKeys.Unitnum = Costco_Wet.UnitNum"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.Unitnum"
    Flag =0
    Expression ="Costco_Wet.LogActivityDate"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
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
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Costco_Wet.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_Wet.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_Wet.Miles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Costco_Wet.milesltd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =140
    Top =36
    Right =1617
    Bottom =920
    Left =-1
    Top =-1
    Right =1459
    Bottom =435
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =147
        Top =-17
        Right =291
        Bottom =356
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
        Name ="Costco_Wet"
        Name =""
    End
End
