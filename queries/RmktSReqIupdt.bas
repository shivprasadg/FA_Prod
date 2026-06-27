Operation =4
Option =0
Where ="(((Clients.clientID)=[forms]![RemarketingClientView]![clientid]) AND ((Rmkt.Invo"
    "icePaidDate) Is Null) AND ((Rmkt.InvoiceRequested)=Yes) AND ((Units.unitstatus)="
    "\"tr\" Or (Units.unitstatus)=\"r\" Or (Units.unitstatus)=\"k\"))"
Begin InputTables
    Name ="Units"
    Name ="Rmkt"
    Name ="Clients"
End
Begin OutputColumns
    Name ="Rmkt.RmkSm"
    Expression ="[forms]![Purchaser Invoice Request]![RSM]"
    Name ="Rmkt.InvoiceRequestDate"
    Expression ="Date()"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Rmkt"
    Expression ="Units.UnitID = Rmkt.UnitRef"
    Flag =1
    LeftTable ="Rmkt"
    RightTable ="Clients"
    Expression ="Rmkt.BuyerRef = Clients.clientID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[Dlr/Resale]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.rmktsaledt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.PInvReq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.RmkSm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.AInvREq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Units].[rmktsaledt]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =152
    Top =224
    Right =1178
    Bottom =755
    Left =-1
    Top =-1
    Right =1008
    Bottom =168
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Rmkt"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
End
