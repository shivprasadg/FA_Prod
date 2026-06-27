SELECT
  Clients.clientCompanyName,
  Clients.clientAddress1,
  Clients.clientAddress2,
  Clients.clientCity,
  Clients.clientState,
  Clients.clientZip,
  Clients.clientMainPhone,
  Clients.clientFax,
  Contacts.contactCreateDate,
  Contacts.LastChangeDt,
  [contactlast] & " " & [contactsuffix] & ", " & [contactfirst] & " " & [contactmiddle] & " " & [contacttitle] AS NM,
  Contacts.contactPrimaryPhone,
  Contacts.contactPhoneExt,
  Contacts.contactPhone2,
  Contacts.contactFax,
  Contacts.contactCell,
  Contacts.contactEmail,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LAddrLine2,
  Locations.LAddrLine3,
  Locations.LCity,
  Locations.LState,
  Locations.LZip,
  Contacts.contactLast,
  Contacts.contactFirst,
  Contacts.contactMiddle,
  Clients.CLIENTTYPE,
  Clients.clienttype,
  IIf(
    IsNull([lastchangedt]),
    "New",
    "Changed"
  ) AS srtr,
  Contacts.ContactObsoletedt,
  Contacts.ContactObsolete
FROM
  Clients
  INNER JOIN (
    Contacts
    LEFT JOIN Locations ON Contacts.cLocationID = Locations.LocationID
  ) ON Clients.clientID = Contacts.clientRef
WHERE
  (
    (
      (Clients.clientCompanyName)<> "FLEET ADVANTAGE"
      And (Clients.clientCompanyName) Not Like "demo*"
    )
    And (
      (Contacts.contactCreateDate)>= forms!call_report_param!startdate
      And (Contacts.contactCreateDate)< forms!call_report_param!enddate
    )
    And (
      (Clients.clienttype)<> "dead"
    )
  )
  Or (
    (
      (Clients.clientCompanyName)<> "FLEET ADVANTAGE"
      And (Clients.clientCompanyName) Not Like "demo*"
    )
    And (
      (Contacts.LastChangeDt)>= forms!call_report_param!startdate
      And (Contacts.LastChangeDt)< forms!call_report_param!enddate
    )
  )
  Or (
    (
      (Contacts.ContactObsoletedt)>= forms!call_report_param!startdate
      And (Contacts.ContactObsoletedt)< forms!call_report_param!enddate
    )
  );
