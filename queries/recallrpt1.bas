Operation =1
Option =0
Where ="(((Clients.ClientShNm)=forms!recallpform!client) And ((Recalls.recallref)=forms!"
    "recallpform!recallno))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Locations"
    Name ="Recalls"
    Name ="RecallCampaign"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="Recalls.recallref"
    Expression ="Clients.clientCompanyName"
    Expression ="Schedule.SchNo"
    Expression ="Units.unitunitnum"
    Expression ="Units.unitvin"
    Expression ="Units.UnitEngineSerial"
    Expression ="Locations.LocationName"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LCity"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="MstrLease.MLNo"
    Expression ="Recalls.RCSent"
    Expression ="MstrLease.Client"
    Expression ="Recalls.RCCont"
    Expression ="RecallCampaign.RecallNo"
    Expression ="Recalls.RCInspDt"
    Expression ="Recalls.RCRepairReq"
    Expression ="Recalls.RCRepDt"
    Expression ="Recalls.[Repair Done]"
    Expression ="Locations.LocationID"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="Units"
    RightTable ="Recalls"
    Expression ="Units.UnitID = Recalls.RCUnitID"
    Flag =1
    LeftTable ="Recalls"
    RightTable ="RecallCampaign"
    Expression ="Recalls.recallref = RecallCampaign.ID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="Units.unitunitnum"
    Flag =0
    Expression ="Locations.LocationID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4185"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Units.UnitEngineSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCCont"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecallCampaign.RecallNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCInspDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCRepairReq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCRepDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.[Repair Done]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.recallref"
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
    Bottom =399
    Left =464
    Top =0
    ColumnsShown =539
    Begin
        Left =-416
        Top =12
        Right =-272
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =-224
        Top =12
        Right =-81
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =-32
        Top =12
        Right =112
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =160
        Top =12
        Right =304
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =352
        Top =12
        Right =557
        Bottom =172
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =368
        Top =189
        Right =511
        Bottom =400
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =615
        Top =-2
        Right =759
        Bottom =173
        Top =0
        Name ="Recalls"
        Name =""
    End
    Begin
        Left =793
        Top =4
        Right =937
        Bottom =148
        Top =0
        Name ="RecallCampaign"
        Name =""
    End
End
