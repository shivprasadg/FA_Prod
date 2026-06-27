Operation =1
Option =0
Begin InputTables
    Name ="Contacts"
    Name ="Locations"
    Name ="lookup_States"
End
Begin OutputColumns
    Expression ="Contacts.*"
    Alias ="ClientId"
    Expression ="Contacts.clientRef"
    Expression ="lookup_States.StateAbbrev"
    Expression ="Locations.LocationID"
    Expression ="Locations.LocationType"
    Expression ="Locations.LocationName"
    Expression ="Locations.Locationsn"
    Expression ="Locations.District"
    Expression ="Locations.LocSubLessee"
    Expression ="Locations.LocationStreet"
    Expression ="Locations.LAddrLine1"
    Expression ="Locations.LAddrLine2"
    Expression ="Locations.LAddrLine3"
    Expression ="Locations.LCity"
    Expression ="Locations.LCounty"
    Expression ="Locations.LState"
    Expression ="Locations.LZip"
    Expression ="Locations.LTaxRate"
    Expression ="Locations.LCountry"
    Expression ="Locations.LPhone"
    Expression ="Locations.LFax"
    Expression ="Locations.LocObsolete"
    Expression ="Locations.LocObsdate"
    Expression ="Locations.LocObsCmt"
    Expression ="Locations.RLocId"
    Expression ="Locations.LocCC"
    Expression ="Locations.LTaxOption"
    Expression ="Locations.LocCostCtr"
    Expression ="Locations.LocABA"
    Expression ="Locations.LocACH"
    Expression ="Locations.LocBKACCT"
    Expression ="Locations.LOCCCC"
    Expression ="Locations.LocObsdate"
    Expression ="Locations.LocObsolete"
End
Begin Joins
    LeftTable ="Contacts"
    RightTable ="Locations"
    Expression ="Contacts.cLocationID = Locations.LocationID"
    Flag =2
    LeftTable ="Locations"
    RightTable ="lookup_States"
    Expression ="Locations.LState = lookup_States.StateAbbrev"
    Flag =2
End
Begin OrderBy
    Expression ="Contacts.contactClassification"
    Flag =0
    Expression ="Contacts.contactLast"
    Flag =0
    Expression ="Contacts.contactFirst"
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
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactOldID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCreateDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Contacts.contactEditDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactClassification"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPrefix"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFirst"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactMiddle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactLast"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactSuffix"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.C-Level"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactSalutation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPrimaryPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPhoneExt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPhone2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCell"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactEmail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactOptOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactInc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCofA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactcounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.oldLocID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.locationtaxrt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.countyRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.cityRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactObsoletedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactRTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.locadddt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LastChangeDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactAsstFnm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactAsstLnm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactAsstE-mail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactAsstPH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactNt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StateAbbrev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LOCCCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsdate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.UpsizeTimeStamp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.Locationsn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.District"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocSubLessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocObsCmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.oldlocid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.RLocId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocCC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocCostCtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocABA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocACH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocBKACCT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.SFDCid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxOption"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.DoNotCall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LAddrLine3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LTaxRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1021"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1022"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.cLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="oldlocid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-328
    Top =280
    Right =1444
    Bottom =926
    Left =-1
    Top =-1
    Right =1748
    Bottom =413
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =294
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =231
        Top =0
        Name ="Locations"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="lookup_States"
        Name =""
    End
End
