Operation =1
Option =0
Where ="(((Clients.clienttype)<>\"Bank\" And (Clients.clienttype)<>\"dead\" And (Clients"
    ".clienttype)<>\"r*\"))"
Begin InputTables
    Name ="Clients"
    Name ="GuarantorsfParent"
End
Begin OutputColumns
    Alias ="Unique ID"
    Expression ="Clients.clientID"
    Alias ="NAME"
    Expression ="Clients.clientCompanyName"
    Alias ="Type"
    Expression ="Clients.clienttype"
    Alias ="Parent"
    Expression ="GuarantorsfParent.clientID"
    Alias ="BILLINGSTREET"
    Expression ="Clients.clientAddress1"
    Alias ="BILLINGCITY"
    Expression ="Clients.clientCity"
    Alias ="BILLINGSTATE"
    Expression ="Clients.clientState"
    Alias ="BILLINGPOSTALCODE"
    Expression ="Clients.clientZip"
    Alias ="BILLINGCOUNTRY"
    Expression ="Clients.clientcountry"
    Alias ="SHIPPINGSTREET"
    Expression ="Null"
    Alias ="SHIPPINGCITY"
    Expression ="Null"
    Alias ="SHIPPINGSTATE"
    Expression ="Null"
    Alias ="SHIPPINGPOSTALCODE"
    Expression ="Null"
    Alias ="SHIPPINGCOUNTRY"
    Expression ="Null"
    Alias ="PHONE"
    Expression ="Clients.clientMainPhone"
    Alias ="Fax"
    Expression ="Clients.clientFax"
    Alias ="Website"
    Expression ="Clients.clientWebsite"
    Alias ="INDUSTRY"
    Expression ="Clients.clientIndustry"
    Alias ="ANNUALREVENUE"
    Expression ="0"
    Alias ="NUMBEROFEMPLOYEES"
    Expression ="0"
    Alias ="DESCRIPTION"
    Expression ="Null"
    Alias ="OWNER"
    Expression ="Clients.LOinitial"
    Alias ="DAYCABS__C"
    Expression ="Clients.fleetDaycabs"
    Alias ="SLEEPERS__C"
    Expression ="Clients.fleetSleepers"
    Alias ="FLEET_SIZE__C"
    Expression ="Clients.fleetsize"
    Alias ="CREDIT_RATING_MOODYS_SP__C"
    Expression ="Clients.creditMoodyRating"
    Alias ="FA_CREDIT_RANK__C"
    Expression ="Clients.creditRank"
    Alias ="FA_CREDIT_SCORE__C"
    Expression ="Clients.creditFAScore"
    Alias ="ACQUISITION__C"
    Expression ="Clients.creditFinance"
    Alias ="MAINTENANCE__C"
    Expression ="Clients.creditmaint"
    Alias ="CYCLE_YEARS__C"
    Expression ="Clients.FleetReplacementCycle"
    Alias ="OBC__C"
    Expression ="Clients.clientOBC"
    Expression ="Clients.ClientGroupID"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="GuarantorsfParent"
    Expression ="Clients.ClientGroupID = GuarantorsfParent.ClientGroupID"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.clientID"
    Flag =0
    Expression ="Clients.clientCompanyName"
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
        dbText "Name" ="Unique ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NAME"
        dbInteger "ColumnWidth" ="3435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BILLINGSTREET"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3420"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="BILLINGCITY"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BILLINGSTATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BILLINGPOSTALCODE"
        dbInteger "ColumnWidth" ="1230"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BILLINGCOUNTRY"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PHONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Website"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="INDUSTRY"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3555"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SHIPPINGSTREET"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SHIPPINGCITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SHIPPINGSTATE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SHIPPINGPOSTALCODE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SHIPPINGCOUNTRY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ANNUALREVENUE"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="NUMBEROFEMPLOYEES"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DESCRIPTION"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="OWNER"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DAYCABS__C"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SLEEPERS__C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FLEET_SIZE__C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CREDIT_RATING_MOODYS_SP__C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FA_CREDIT_RANK__C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ACQUISITION__C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MAINTENANCE__C"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="CYCLE_YEARS__C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC__C"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1185"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FA_CREDIT_SCORE__C"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[client group]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Parent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1262
    Bottom =793
    Left =-1
    Top =-1
    Right =1230
    Bottom =336
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =270
        Bottom =334
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =443
        Top =27
        Right =587
        Bottom =171
        Top =0
        Name ="GuarantorsfParent"
        Name =""
    End
End
