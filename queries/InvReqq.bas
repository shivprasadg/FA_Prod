Operation =1
Option =0
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="Clients.clientAddress1"
    Expression ="Clients.clientAddress2"
    Expression ="Clients.clientCity"
    Expression ="Clients.clientState"
    Expression ="Clients.clientZip"
    Expression ="Clients.[Dlr/Resale]"
    Expression ="Units.unitvin"
    Expression ="Units.unityr"
    Expression ="Units.unitmake"
    Expression ="Units.unittype"
    Expression ="Clients.clientID"
End
Begin Joins
    LeftTable ="Rmkt"
    RightTable ="Clients"
    Expression ="Rmkt.BuyerRef = Clients.clientID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
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
Begin
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unityr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitmake"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1544
    Bottom =772
    Left =-1
    Top =-1
    Right =1512
    Bottom =415
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-3
        Top =1
        Right =141
        Bottom =145
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =142
        Top =16
        Right =286
        Bottom =160
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =328
        Top =17
        Right =566
        Bottom =268
        Top =0
        Name ="Clients"
        Name =""
    End
End
