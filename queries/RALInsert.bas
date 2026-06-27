Operation =3
Name ="RALTemp1"
Option =0
Begin InputTables
    Name ="RALSource"
End
Begin OutputColumns
    Name ="clientID"
    Expression ="RALSource.clientID"
    Name ="SchID"
    Expression ="RALSource.SchID"
    Name ="RemitAddressID"
    Expression ="RALSource.RemitAddressID"
    Alias ="Letterdate"
    Name ="Letterdate"
    Expression ="Format(Date(),\"mmmm dd\"\", \"\"yyyy\") & \"<br />\""
    Alias ="Salutation"
    Name ="Salutation"
    Expression ="\"Gentlemen:<br />\""
    Name ="ScheduleDate"
    Expression ="RALSource.ScheduleDate"
    Name ="Schedule"
    Expression ="RALSource.Schedule"
    Name ="MLNo"
    Expression ="RALSource.MLNo"
    Name ="MLDate"
    Expression ="RALSource.MLDate"
    Name ="Lessee"
    Expression ="RALSource.Lessee"
    Name ="LesseeAddress"
    Expression ="RALSource.LesseeAddress"
    Name ="BankRemitAddress"
    Expression ="RALSource.BankRemitAddress"
    Name ="Paragraph1"
    Expression ="RALSource.Paragraph1"
    Alias ="Paragraph2"
    Name ="Paragraph2"
    Expression ="[Paragraph2a] & [Paragraph2b]"
    Alias ="Paragraph3"
    Name ="Paragraph3"
    Expression ="[Paragraph3a] & [Paragraph3b] & [Paragraph3c]"
    Name ="Paragraph4"
    Expression ="RALSource.Paragraph4"
    Name ="Paragraph5"
    Expression ="RALSource.Paragraph5"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="RALSource.Paragraph2a"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Paragraph5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.LesseeAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Paragraph1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.BankRemitAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Paragraph2b"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Paragraph3a"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Paragraph3b"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Paragraph3c"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Paragraph4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.RemitAddressID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.ScheduleDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.MLDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RALSource.Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paragraph3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salutation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Letterdate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =84
    Top =-3
    Right =1655
    Bottom =943
    Left =-1
    Top =-1
    Right =1553
    Bottom =601
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =92
        Top =54
        Right =236
        Bottom =466
        Top =0
        Name ="RALSource"
        Name =""
    End
End
