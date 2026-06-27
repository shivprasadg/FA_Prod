Operation =4
Option =0
Begin InputTables
    Name ="Zipcty1"
End
Begin OutputColumns
    Name ="Zipcty1.State"
    Expression ="Mid([Field1],24,2)"
    Name ="Zipcty1.FIPS"
    Expression ="Mid([Field1],26,3)"
    Name ="Zipcty1.County"
    Expression ="Mid([Field1],29,25)"
    Name ="Zipcty1.Plus4"
    Expression ="Mid([Field1],16,4)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Zipcty1.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zipcty1.FIPS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zipcty1.County"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zipcty1.Plus4"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1477
    Bottom =1026
    Left =-1
    Top =-1
    Right =1459
    Bottom =709
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =83
        Top =26
        Right =227
        Bottom =170
        Top =0
        Name ="Zipcty1"
        Name =""
    End
End
