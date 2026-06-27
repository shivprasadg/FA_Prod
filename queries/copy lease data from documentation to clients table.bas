Operation =4
Option =0
Where ="((Not (Clients.docMLAContactREF) Is Null)) Or ((Not (Clients.docMLASigRef) Is Nu"
    "ll)) Or ((Not (Clients.docMLASigTitle) Is Null)) Or ((Not (Clients.docGuarSigREF"
    ") Is Null)) Or ((Not (Clients.docGuarSigrTitle) Is Null)) Or ((Not (Clients.docL"
    "easeID) Is Null)) Or ((Not (Clients.docNameVer) Is Null)) Or ((Not (Clients.docC"
    "GS) Is Null)) Or ((Not (Clients.docTaxID) Is Null))"
Begin InputTables
    Name ="Clients"
    Name ="Deals"
    Name ="Documentation"
End
Begin OutputColumns
    Name ="Clients.docMLAContactREF"
    Expression ="[Documentation].[docMLAContactREF]"
    Name ="Clients.docMLASigRef"
    Expression ="[documentation].[docMLASigRef]"
    Name ="Clients.docMLASigTitle"
    Expression ="[documentation].[docMLASigTitle]"
    Name ="Clients.docGuarSigREF"
    Expression ="[documentation].[docGuarSigREF]"
    Name ="Clients.docGuarSigrTitle"
    Expression ="[documentation].[docGuarSigrTitle]"
    Name ="Clients.docLeaseID"
    Expression ="[documentation].[docLeaseID]"
    Name ="Clients.docNameVer"
    Expression ="[documentation].[docNameVer]"
    Name ="Clients.docCGS"
    Expression ="[documentation].[docCGS]"
    Name ="Clients.docTaxID"
    Expression ="[documentation].[docTaxID]"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
    Flag =1
    LeftTable ="Deals"
    RightTable ="Documentation"
    Expression ="Deals.dealID = Documentation.dealRef"
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
        dbText "Name" ="Deals.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docMLAContactREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docMLASigRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docMLASigTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docGuarSigREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docGuarSigrTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docLeaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docNameVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docCGS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Documentation.docTaxID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="docMLAContactREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docMLAContactREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docMLASigRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docMLASigTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docGuarSigREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docGuarSigrTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docLeaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docNameVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docCGS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.docTaxID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
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
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1152
    Bottom =678
    Left =-1
    Top =-1
    Right =1120
    Bottom =335
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =231
        Top =26
        Right =375
        Bottom =170
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =444
        Top =20
        Right =715
        Bottom =269
        Top =0
        Name ="Documentation"
        Name =""
    End
End
