Operation =1
Option =0
Where ="(((Clients.ClientGroupID) Is Not Null))"
Begin InputTables
    Name ="Clients"
End
Begin OutputColumns
    Expression ="Clients.ClientGroupID"
    Expression ="Clients.ClientShNm"
    Alias ="Flt"
    Expression ="IIf([fleetDaycabs]+[fleetSleepers]<200,-0.1,0.15+0.05*(-Int(-(([fleetdaycabs]+[f"
        "leetsleepers]-200)/500))))"
    Alias ="DCMI"
    Expression ="IIf([FLEETDAYCABSMPY]=0,0,IIf([FLEETDAYCABSMPY]<60000,-0.4,IIf([FLEETDAYCABSMPY]"
        "<70000,-0.2,IIf([FLEETDAYCABSMPY]<80000,-0.05,0))))"
    Alias ="SlpMI"
    Expression ="IIf([fleetSleepersMPY]=0,0,IIf([fleetSleepersMPY]<70000,-0.4,IIf([fleetSleepersM"
        "PY]<80000,-0.2,IIf([fleetSleepersMPY]<90000,-0.1,0))))"
    Alias ="AGRT"
    Expression ="IIf([creditrank]<>\"tb\" And [creditMoodyRating]=\"tbd\",0,IIf([creditmoodyratin"
        "g]=\"aaa\",0.4,IIf(Left([creditmoodyrating],2)=\"aa\",0.35,IIf(Left([creditmoody"
        "rating],1)=\"a\",0.3,IIf(Left([creditmoodyrating],3)=\"baa\",0.25,IIf(Left([cred"
        "itmoodyrating],2)=\"BAA\",0.25,IIf(Left([creditmoodyrating],3)=\"BBB\",0.25,IIf("
        "Left([creditmoodyrating],2)=\"BA\",0.15,IIf(Left([creditmoodyrating],2)=\"BB\",0"
        ".15,-0.35)))))))))"
    Alias ="FART"
    Expression ="IIf([CREDITRANK]=\"1\",0.4,IIf([CREDITRANK]=\"2\",0.35,IIf([CREDITRANK]=\"3\",0."
        "3,IIf([CREDITRANK]=\"4\",0.25,IIf([CREDITRANK]=\"5\",0.15,-0.35)))))"
    Alias ="AQ"
    Expression ="IIf([creditfinance]=\"purchase\",0.15,IIf([creditfinance]=\"lease\",0.1,IIf([cre"
        "ditfinance]=\"FULL service\",-0.2,0)))"
    Alias ="Maint"
    Expression ="IIf([creditfinance]=\"full service\",0,IIf([creditmaint]=\"in-house\",0.15,IIf(["
        "creditmaint]=\"outsource\" And [CREDITFINANCE]<>\"FULL Service\",0.05,0)))"
    Alias ="Cyc"
    Expression ="IIf(IsNull([FleetReplacementCycle]),0,IIf([FleetReplacementCycle]>3.99,0.15,0))"
    Alias ="OBC"
    Expression ="IIf([clientobc]=\"None\" Or [clientobc]=\"tbd\",0,0.1)"
End
Begin OrderBy
    Expression ="Clients.ClientGroupID"
    Flag =0
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
        dbText "Name" ="Flt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
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
        dbText "Name" ="AGRT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Maint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cyc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[CLIENT GROUP]"
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
    Bottom =524
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
