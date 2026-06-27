Operation =1
Option =0
Begin InputTables
    Name ="UnitsTypesAccepted_ByYear"
    Name ="UnitsTypesNewBusiness_ByYear"
    Name ="UnitsTypesEOL_ByYear"
End
Begin OutputColumns
    Expression ="UnitsTypesAccepted_ByYear.Accepted"
    Expression ="UnitsTypesAccepted_ByYear.UnitType"
    Alias ="SumOfUnitAccepted"
    Expression ="Sum(UnitsTypesAccepted_ByYear.UnitAccepted)"
    Expression ="UnitsTypesNewBusiness_ByYear.NewBusiness"
    Expression ="UnitsTypesEOL_ByYear.UnitEOL"
End
Begin Joins
    LeftTable ="UnitsTypesAccepted_ByYear"
    RightTable ="UnitsTypesNewBusiness_ByYear"
    Expression ="UnitsTypesAccepted_ByYear.UnitType = UnitsTypesNewBusiness_ByYear.UnitType"
    Flag =2
    LeftTable ="UnitsTypesAccepted_ByYear"
    RightTable ="UnitsTypesNewBusiness_ByYear"
    Expression ="UnitsTypesAccepted_ByYear.Accepted = UnitsTypesNewBusiness_ByYear.NewClient"
    Flag =2
    LeftTable ="UnitsTypesAccepted_ByYear"
    RightTable ="UnitsTypesEOL_ByYear"
    Expression ="UnitsTypesAccepted_ByYear.Accepted = UnitsTypesEOL_ByYear.EOL"
    Flag =2
    LeftTable ="UnitsTypesAccepted_ByYear"
    RightTable ="UnitsTypesEOL_ByYear"
    Expression ="UnitsTypesAccepted_ByYear.UnitType = UnitsTypesEOL_ByYear.UnitType"
    Flag =2
End
Begin Groups
    Expression ="UnitsTypesAccepted_ByYear.Accepted"
    GroupLevel =0
    Expression ="UnitsTypesAccepted_ByYear.UnitType"
    GroupLevel =0
    Expression ="UnitsTypesNewBusiness_ByYear.NewBusiness"
    GroupLevel =0
    Expression ="UnitsTypesEOL_ByYear.UnitEOL"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([Query3].[Accepted]=2017)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SumOfUnitAccepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsTypesAccepted_ByYear.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsTypesAccepted_ByYear.Accepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsTypesNewBusiness_ByYear.NewBusiness"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsTypesEOL_ByYear.UnitEOL"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =147
    Top =37
    Right =1454
    Bottom =966
    Left =-1
    Top =-1
    Right =1289
    Bottom =652
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =109
        Top =47
        Right =253
        Bottom =228
        Top =0
        Name ="UnitsTypesAccepted_ByYear"
        Name =""
    End
    Begin
        Left =500
        Top =290
        Right =644
        Bottom =434
        Top =0
        Name ="UnitsTypesNewBusiness_ByYear"
        Name =""
    End
    Begin
        Left =498
        Top =141
        Right =642
        Bottom =285
        Top =0
        Name ="UnitsTypesEOL_ByYear"
        Name =""
    End
End
