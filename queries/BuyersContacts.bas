Operation =1
Option =0
Begin InputTables
    Name ="Contacts"
    Name ="Buyers"
End
Begin OutputColumns
    Expression ="Contacts.*"
    Alias ="Contact"
    Expression ="[contactFirst] & \" \" & [contactLast]"
End
Begin Joins
    LeftTable ="Buyers"
    RightTable ="Contacts"
    Expression ="Buyers.clientID = Contacts.clientRef"
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
        dbText "Name" ="Contacts.contactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactClassification"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPhoneExt"
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
        dbText "Name" ="Contacts.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactEditDate"
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
        dbText "Name" ="Contacts.oldLocID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactOldID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactCreateDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPrefix"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactAddress2"
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
        dbText "Name" ="Contacts.LocationName"
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
        dbText "Name" ="Contacts.contactRTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.C-Level"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactPrimaryPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactSalutation"
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
        dbText "Name" ="Contacts.DoNotCall"
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
        dbText "Name" ="Contacts.LocSlprs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.cityRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.locSplrMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LocTrucks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LocDCs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LocTrkMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactObsoletedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LocDCMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.LocIdref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactAsstE-mail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.UpsizeTimeStamp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.locadddt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactAsstLnm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactMorpt"
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
        dbText "Name" ="Contacts.ContactAsstPH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.ContactNt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contact"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1528
    Bottom =1013
    Left =-1
    Top =-1
    Right =1496
    Bottom =655
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =263
        Top =70
        Right =404
        Bottom =463
        Top =0
        Name ="Contacts"
        Name =""
    End
    Begin
        Left =48
        Top =38
        Right =192
        Bottom =182
        Top =0
        Name ="Buyers"
        Name =""
    End
End
