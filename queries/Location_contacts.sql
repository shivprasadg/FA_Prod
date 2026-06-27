SELECT
  Contacts.*,
  Contacts.clientRef AS ClientId,
  lookup_States.StateAbbrev,
  Locations.LocationID,
  Locations.LocationType,
  Locations.LocationName,
  Locations.Locationsn,
  Locations.District,
  Locations.LocSubLessee,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LAddrLine2,
  Locations.LAddrLine3,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  Locations.LTaxRate,
  Locations.LCountry,
  Locations.LPhone,
  Locations.LFax,
  Locations.LocObsolete,
  Locations.LocObsdate,
  Locations.LocObsCmt,
  Locations.RLocId,
  Locations.LocCC,
  Locations.LTaxOption,
  Locations.LocCostCtr,
  Locations.LocABA,
  Locations.LocACH,
  Locations.LocBKACCT,
  Locations.LOCCCC,
  Locations.LocObsdate,
  Locations.LocObsolete
FROM
  (
    Contacts
    LEFT JOIN Locations ON Contacts.cLocationID = Locations.LocationID
  )
  LEFT JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
ORDER BY
  Contacts.contactClassification,
  Contacts.contactLast,
  Contacts.contactFirst;
