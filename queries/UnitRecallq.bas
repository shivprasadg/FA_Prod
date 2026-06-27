Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
    Name ="Units"
    Name ="Recalls"
    Name ="Locations"
    Name ="RecallCampaign"
    Name ="Contacts"
End
Begin OutputColumns
    Alias ="Company"
    Expression ="Clients.clientCompanyName"
    Expression ="RecallCampaign.RecallNo"
    Expression ="RecallCampaign.Type"
    Alias ="Descsription"
    Expression ="RecallCampaign.RecallDesc"
    Alias ="Recall Dt"
    Expression ="RecallCampaign.RecallDt"
    Alias ="RA Received"
    Expression ="RecallCampaign.RecallRec"
    Alias ="Unit No"
    Expression ="Units.UnitUnitNum"
    Alias ="VIIN"
    Expression ="Units.unitvin"
    Alias ="Eng Serioa NO"
    Expression ="Units.UnitEngineSerial"
    Expression ="Locations.LocationName"
    Alias ="Add1"
    Expression ="Locations.LocationStreet"
    Alias ="Add2"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LCity"
    Alias ="St"
    Expression ="Locations.LState"
    Alias ="ZiP"
    Expression ="Locations.LZip"
    Expression ="Schedule.SchNo"
    Alias ="Inspected (date)"
    Expression ="Null"
    Alias ="Repair Needed (Y/N)"
    Expression ="Null"
    Alias ="Repair Made(dt)"
    Expression ="Null"
    Expression ="Contacts.contactFullName"
    Expression ="Contacts.contactPrimaryPhone"
    Expression ="Contacts.contactEmail"
    Expression ="Units.UnitID"
    Expression ="Recalls.RCSent"
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
    RightTable ="Recalls"
    Expression ="Units.UnitID = Recalls.RCUnitID"
    Flag =2
    LeftTable ="Units"
    RightTable ="Locations"
    Expression ="Units.UnitGarageLocationRef = Locations.LocationID"
    Flag =1
    LeftTable ="Recalls"
    RightTable ="RecallCampaign"
    Expression ="Recalls.recallref = RecallCampaign.ID"
    Flag =2
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
    Expression ="RecallCampaign.RecallDt"
    Flag =1
    Expression ="Locations.LocationName"
    Flag =0
    Expression ="MstrLease.Client"
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
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecallCampaign.RecallNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecallCampaign.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Inspected (date)"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repair Needed (Y/N)"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Company"
        dbInteger "ColumnWidth" ="4335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Descsription"
        dbInteger "ColumnWidth" ="5565"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recall Dt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RA Received"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIIN"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Eng Serioa NO"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Add1"
        dbInteger "ColumnWidth" ="4335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Add2"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ZiP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repair Made(dt)"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPrimaryPhone"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactEmail"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Recalls.RCSent"
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
    Bottom =432
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
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =817
        Top =221
        Right =961
        Bottom =467
        Top =0
        Name ="Recalls"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =555
        Top =184
        Right =699
        Bottom =328
        Top =0
        Name ="RecallCampaign"
        Name =""
    End
    Begin
        Left =979
        Top =216
        Right =1244
        Bottom =360
        Top =0
        Name ="Contacts"
        Name =""
    End
End
