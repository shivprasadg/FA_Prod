Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Contacts"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Units.UnitRegLocationRef"
    Expression ="Contacts.oldLocID"
    Expression ="Contacts.contactID"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Contacts"
    Expression ="Units.UnitRegLocationRef = Contacts.oldLocID"
    Flag =1
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Contacts.oldLocID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-1
    Top =4
    Right =977
    Bottom =642
    Left =-1
    Top =-1
    Right =946
    Bottom =335
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =291
        Bottom =262
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =339
        Top =12
        Right =566
        Bottom =286
        Top =0
        Name ="Contacts"
        Name =""
    End
End
