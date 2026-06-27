Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="RecallCampaign"
    Name ="Units"
    Name ="Locations"
    Name ="Recalls"
    Name ="Contacts"
End
Begin OutputColumns
    Alias ="Company"
    Expression ="Clients.clientCompanyName"
    Alias ="RecallNum"
    Expression ="RecallCampaign.RecallNo"
    Expression ="RecallCampaign.Type"
    Alias ="Description"
    Expression ="RecallCampaign.RecallDesc"
    Alias ="NoticeDate"
    Expression ="RecallCampaign.RecallDt"
    Alias ="UnitNumber"
    Expression ="Units.UnitUnitNum"
    Expression ="Units.UnitVIN"
    Expression ="Locations.LocationName"
    Alias ="ActionNeeded"
    Expression ="Null"
    Alias ="ActionCompleted"
    Expression ="Recalls.RCRepDt"
    Alias ="EngSerialNo"
    Expression ="Units.UnitEngineSerial"
    Expression ="Clients.ClientID"
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
    LeftTable ="RecallCampaign"
    RightTable ="Recalls"
    Expression ="RecallCampaign.ID = Recalls.recallref"
    Flag =1
    LeftTable ="Recalls"
    RightTable ="Contacts"
    Expression ="Recalls.RCCont = Contacts.contactID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
End
Begin OrderBy
    Expression ="Locations.LocationName"
    Flag =0
    Expression ="Clients.ClientID"
    Flag =0
    Expression ="Left([unitvin],6)"
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
        dbText "Name" ="Locations.LocationName"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecallCampaign.Type"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Company"
        dbInteger "ColumnWidth" ="3240"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="3675"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecallNum"
        dbInteger "ColumnWidth" ="1725"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActionNeeded"
        dbInteger "ColumnWidth" ="1920"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoticeDate"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActionCompleted"
        dbInteger "ColumnWidth" ="1275"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EngSerialNo"
        dbInteger "ColumnWidth" ="1635"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNumber"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVin"
        dbInteger "ColumnWidth" ="2115"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =56
    Top =354
    Right =1314
    Bottom =1000
    Left =-1
    Top =-1
    Right =1226
    Bottom =141
    Left =0
    Top =96
    ColumnsShown =539
    Begin
        Left =16
        Top =-84
        Right =160
        Bottom =60
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =186
        Top =-85
        Right =330
        Bottom =59
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =347
        Top =-84
        Right =491
        Bottom =60
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =516
        Top =-82
        Right =660
        Bottom =62
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =1060
        Top =-74
        Right =1225
        Bottom =68
        Top =0
        Name ="RecallCampaign"
        Name =""
    End
    Begin
        Left =694
        Top =-81
        Right =838
        Bottom =63
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =875
        Top =92
        Right =1034
        Bottom =236
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =871
        Top =-82
        Right =1033
        Bottom =62
        Top =0
        Name ="Recalls"
        Name =""
    End
    Begin
        Left =1074
        Top =109
        Right =1240
        Bottom =253
        Top =0
        Name ="Contacts"
        Name =""
    End
End
