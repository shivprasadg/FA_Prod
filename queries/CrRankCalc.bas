Operation =1
Option =0
Where ="(((Clients.clientType)<>\"vendor\" And (Clients.clientType)<>\"bank\" And (Clien"
    "ts.clientType)<>\"fa\"))"
Begin InputTables
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.clientID"
    Expression ="Clients.ClientShNm"
    Expression ="Clients.clientType"
    Alias ="Fleet"
    Expression ="IIf([fleetdaycabs]+[fleetsleepers]>=400,0.25,IIf([fleetdaycabs]+[fleetsleepers]>"
        "=200,0.15,0))"
    Alias ="DCMI"
    Expression ="IIf([FLEETDAYCABS]=0,0,IIf([FLEETDAYCABSMPY]=0 Or [fleetdaycabsmpy]>=80000,0,IIf"
        "([FLEETDAYCABSMPY]>70000,-0.1,IIf([FLEETDAYCABSMPY]>60000,-0.2,-0.3))))"
    Alias ="SlpMI"
    Expression ="IIf([FLEETSLEEPERS]=0,0,IIf([FLEETSLEEPERSMPY]=0 Or [FLEETSLEEPERSMPY]>=10000,0,"
        "-0.3))"
    Alias ="Cycle"
    Expression ="IIf(IsNull([FleetReplacementCycle]),0,IIf([FleetReplacementCycle]>4.99,0.15,0))"
    Alias ="OldOBC"
    Expression ="IIf([clientobc]=\"none\",0.05,IIf([clientobc]=\"peoplenet\",0.1,0))"
    Alias ="Moody"
    Expression ="IIf([creditmoodyrating]=\"aaa\",0.4,IIf(Left([creditmoodyrating],2)=\"aa\",0.35,"
        "IIf(Left([creditmoodyrating],1)=\"a\",0.3,IIf(Left([creditmoodyrating],3)=\"baa\""
        ",0.25,IIf(Left([creditmoodyrating],2)=\"BAA\",0.25,IIf(Left([creditmoodyrating],"
        "3)=\"BBB\",0.25,IIf(Left([creditmoodyrating],2)=\"BA\",0.15,IIf(Left([creditmood"
        "yrating],2)=\"BB\",0.15,0))))))))"
    Alias ="CreditBank"
    Expression ="IIf([CREDITRANK]=\"1\",0.4,IIf([CREDITRANK]=\"2\",0.35,IIf([CREDITRANK]=\"3\",0."
        "3,IIf([CREDITRANK]=\"4\",0.25,IIf([CREDITRANK]=\"5\",0.15,0)))))"
    Alias ="Acquisition"
    Expression ="IIf([creditfinance]=\"purchase\",0.15,IIf([creditfinance]=\"lease\",0.1,IIf([cre"
        "ditfinance]=\"FULL\",-0.1,0)))"
    Alias ="Maint"
    Expression ="IIf([creditmaint]=\"in-house\",0.15,IIf([creditmaint]=\"outsource\" And [CREDITF"
        "INANCE]<>\"FULL\",0.05,0))"
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
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
        dbText "Name" ="Fleet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCMI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SlpMI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OldOBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Moody"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CreditBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Acquisition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Maint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =158
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
End
