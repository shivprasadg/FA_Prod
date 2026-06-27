Operation =1
Option =0
Begin InputTables
    Name ="RALSource"
End
Begin OutputColumns
    Expression ="RALSource.clientID"
    Expression ="RALSource.SchID"
    Expression ="RALSource.RemitAddressID"
    Alias ="Letterdate"
    Expression ="Date()"
    Alias ="Salutation"
    Expression ="\"Gentlemen:\""
    Expression ="RALSource.ScheduleDate"
    Expression ="RALSource.Schedule"
    Expression ="RALSource.MLNo"
    Expression ="RALSource.MLDate"
    Expression ="RALSource.Lessee"
    Expression ="RALSource.LesseeAddress"
    Expression ="RALSource.BankRemitAddress"
    Expression ="RALSource.Paragraph1"
    Alias ="Paragraph2"
    Expression ="[Paragraph2a] & [Paragraph2b]"
    Alias ="Paragraph3"
    Expression ="[Paragraph3a] & [Paragraph3b] & [Paragraph3c]"
    Expression ="RALSource.Paragraph4"
    Expression ="RALSource.Paragraph5"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
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
        dbText "Name" ="RALSource.Paragraph4"
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
    Left =222
    Top =73
    Right =1793
    Bottom =1019
    Left =-1
    Top =-1
    Right =1553
    Bottom =567
    Left =0
    Top =0
    ColumnsShown =539
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
