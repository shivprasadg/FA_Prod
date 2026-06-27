Operation =1
Option =0
Where ="(((Deals.dealShortDesc)<>\"Tracking Units\") AND ((Left([dealshortdesc],3))<>\"d"
    "up\") AND ((Left([dealshortdesc],8))<>\"new deal\"))"
Begin InputTables
    Name ="Clients"
    Name ="Deals"
    Name ="lookup_transTypes"
End
Begin OutputColumns
    Alias ="ID"
    Expression ="Deals.dealID"
    Alias ="IsDeleted"
    Expression ="Null"
    Alias ="AccountId"
    Expression ="Deals.clientref"
    Alias ="IsPrivate"
    Expression ="0"
    Alias ="Name"
    Expression ="Deals.dealShortDesc"
    Alias ="Description"
    Expression ="Deals.dealShortDesc"
    Alias ="StageNm"
    Expression ="IIf(IsNull([dealoutcome]),IIf(IsNull([deallpafsent]),IIf(IsNull([dealLpafdue]),["
        "transType],\"Proposal Due\"),\"Proposal Sent\"),[dealoutcome])"
    Alias ="SortOrder"
    Expression ="\"tbd\""
    Alias ="Amount"
    Expression ="IIf(IsNull([dealgrandtotal]),0,[dealgrandtotal])"
    Alias ="Probability"
    Expression ="0"
    Alias ="ExpectedRevenue"
    Expression ="0"
    Alias ="TotalOpportunityQuantity"
    Expression ="Deals.dealunitqty"
    Alias ="CloseDate"
    Expression ="Null"
    Alias ="Type"
    Expression ="IIf(([clienttype]<>\"client\" And [clienttype]<>\"guarantor\") Or IsNull([dealen"
        "tereddate]) Or [dealentereddate]<[clientbecamedate],\"New Business\",\"Existing "
        "Business\")"
    Alias ="NextStep"
    Expression ="IIf(IsNull([dealoutcome]),\"Awarded\",Null)"
    Alias ="LeadSource"
    Expression ="Null"
    Alias ="Is Closed"
    Expression ="IIf(IsNull([dealoutcome]),0,1)"
    Alias ="Is Won"
    Expression ="IIf([dealoutcome]=\"awarded\",1,0)"
    Alias ="ForecastCategory"
    Expression ="IIf(IsNull([dealoutcome]),\"Pipeline\",\"Closed\")"
    Alias ="ForecastCategoryName"
    Expression ="IIf(IsNull([dealoutcome]),\"Pipeline\",\"Closed\")"
    Alias ="CampaignId"
    Expression ="Null"
    Alias ="HasOpportunityLineItem"
    Expression ="Null"
    Alias ="Pricebook2Id"
    Expression ="Null"
    Alias ="OwnerId"
    Expression ="Clients.loinitial"
    Alias ="Createdate"
    Expression ="Deals.dealEnteredDate"
    Alias ="CreatedById"
    Expression ="Clients.loinitial"
    Alias ="dealLPAFDueModifiedOnifiedDate"
    Expression ="Null"
    Alias ="ModifiedOnifiedById"
    Expression ="Null"
    Alias ="SystemModstamp"
    Expression ="Null"
    Alias ="LastActivityDate"
    Expression ="Null"
    Alias ="LastStageChangeDate"
    Expression ="Null"
    Expression ="Clients.creditFYE"
    Alias ="FiscalYear"
    Expression ="Null"
    Alias ="FiscalQuarter"
    Expression ="Null"
    Alias ="PrimaryPartnerAccountId"
    Expression ="Null"
    Alias ="Stage_Changed_C"
    Expression ="Null"
    Alias ="FAID__c"
    Expression ="Deals.faid"
End
Begin Joins
    LeftTable ="Deals"
    RightTable ="lookup_transTypes"
    Expression ="Deals.transTypeRef = lookup_transTypes.TransTypeID"
    Flag =2
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
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
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="615"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="IsDeleted"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="AccountId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1350"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="IsPrivate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="8040"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5835"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="StageNm"
        dbInteger "ColumnWidth" ="1260"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SortOrder"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Amount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="Probability"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="ExpectedRevenue"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="TotalOpportunityQuantity"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2835"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="CloseDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Type"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="NextStep"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeadSource"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Is Closed"
        dbInteger "ColumnWidth" ="1245"
        dbInteger "ColumnOrder" ="17"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Is Won"
        dbInteger "ColumnWidth" ="1005"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ForecastCategory"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="19"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ForecastCategoryName"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="20"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CampaignId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="HasOpportunityLineItem"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Pricebook2Id"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="OwnerId"
        dbInteger "ColumnWidth" ="1230"
        dbInteger "ColumnOrder" ="24"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Createdate"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="CreatedById"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="SystemModstamp"
        dbInteger "ColumnWidth" ="1650"
        dbInteger "ColumnOrder" ="28"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastActivityDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="LastStageChangeDate"
        dbInteger "ColumnWidth" ="1650"
        dbInteger "ColumnOrder" ="30"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFYE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FiscalYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FiscalQuarter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrimaryPartnerAccountId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stage_Changed_C"
        dbInteger "ColumnWidth" ="1365"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID__c"
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
    Bottom =301
    Left =308
    Top =0
    ColumnsShown =539
    Begin
        Left =21
        Top =14
        Right =250
        Bottom =217
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =271
        Top =45
        Right =492
        Bottom =224
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =526
        Top =75
        Right =670
        Bottom =219
        Top =0
        Name ="lookup_transTypes"
        Name =""
    End
End
