Operation =1
Option =0
Where ="(((RecallCampaign.ID)=forms!recallpform!recallno) And ((Clients.ClientShNm)=form"
    "s!recallpform!client) And ((Locations.LocationID) Like IIf(Nz(forms!recallpform!"
    "clientcity),forms!recallpform!clientcity,\"*\")))"
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
    Alias ="No"
    Expression ="RecallCampaign.RecallNo"
    Expression ="RecallCampaign.Type"
    Alias ="Description"
    Expression ="RecallCampaign.RecallDesc"
    Alias ="Notice Dt"
    Expression ="RecallCampaign.RecallDt"
    Alias ="Unit No"
    Expression ="Units.UnitUnitNum"
    Alias ="VIIN"
    Expression ="Units.unitvin"
    Expression ="Locations.LocationName"
    Alias ="Action Needed?"
    Expression ="Null"
    Alias ="Action Completed( Dt)"
    Expression ="Recalls.RCRepDt"
    Alias ="Eng Serial No"
    Expression ="Units.UnitEngineSerial"
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
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="RecallCampaign.Type"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Company"
        dbInteger "ColumnWidth" ="3240"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Unit No"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="VIIN"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Eng Serial No"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Description"
        dbInteger "ColumnWidth" ="3675"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="NO"
        dbInteger "ColumnWidth" ="1725"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notice Dt"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Action Needed?"
        dbInteger "ColumnWidth" ="1920"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Action Completed( Dt)"
        dbInteger "ColumnWidth" ="1275"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =243
    Left =1547
    Top =0
    ColumnsShown =539
    Begin
        Left =-1531
        Top =12
        Right =-1387
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =-1361
        Top =11
        Right =-1217
        Bottom =155
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =-1200
        Top =12
        Right =-1056
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =-1031
        Top =14
        Right =-887
        Bottom =158
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =-487
        Top =22
        Right =-322
        Bottom =164
        Top =0
        Name ="RecallCampaign"
        Name =""
    End
    Begin
        Left =-853
        Top =15
        Right =-709
        Bottom =159
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =-672
        Top =188
        Right =-513
        Bottom =332
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =-676
        Top =14
        Right =-514
        Bottom =158
        Top =0
        Name ="Recalls"
        Name =""
    End
    Begin
        Left =-473
        Top =205
        Right =-307
        Bottom =349
        Top =0
        Name ="Contacts"
        Name =""
    End
End
