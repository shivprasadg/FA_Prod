Operation =1
Option =0
Begin InputTables
    Name ="TrailerMilesHours"
End
Begin OutputColumns
    Expression ="TrailerMilesHours.UnitID"
    Expression ="TrailerMilesHours.UnitNumber"
    Expression ="TrailerMilesHours.LogActivityDate"
    Alias ="MilesCode"
    Expression ="Switch([Miles]=0,0,[Miles]>=1 And [Miles]<35000,1,[Miles]<0,16,[Miles]>=35000,14"
        ")"
    Alias ="MilesCodeDesc"
    Expression ="Switch([MilesCode]=14,\"Excess Miles\",[MilesCode]=16,\"Neg Miles\",[MilesCode]="
        "0,\"No Miles\",[MilesCode]=1,\"Good Miles\")"
    Alias ="HoursCode"
    Expression ="Switch([Hours]>1000,22,[Hours]<0,33,[Hours]=0,0,[Hours]>=1 And [Hours]<1000,3)"
    Alias ="HoursCodeDesc"
    Expression ="Switch([HoursCode]=22,\"Excess Hours\",[HoursCode]=33,\"Neg Hours\",[HoursCode]="
        "0,\"No Hours\",[HoursCode]=3,\"Good Hours\")"
    Alias ="DataStatus"
    Expression ="IIf([MilesCode]+[HoursCode]<5,1,[MilesCode]+[HoursCode])"
    Alias ="DataStatusDesc"
    Expression ="Switch([MilesCode]+[HoursCode]=0,\"Reported - No Update\",[MilesCode]+[HoursCode"
        "]=1,\"Reported - No Hours\",[MilesCode]+[HoursCode]=3,\"Reported - No Miles\",[M"
        "ilesCode]+[HoursCode]=4,\"Reporting\",[MilesCode]+[HoursCode]>4,[MilesCodeDesc] "
        "& \"-\" & [HoursCodeDesc])"
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
        dbText "Name" ="TrailerMilesHours.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HoursCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DataStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MilesCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HoursCodeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MilesCodeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DataStatusDesc"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =26
    Top =-3
    Right =1498
    Bottom =995
    Left =-1
    Top =-1
    Right =1440
    Bottom =680
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =154
        Top =103
        Right =641
        Bottom =487
        Top =0
        Name ="TrailerMilesHours"
        Name =""
    End
End
