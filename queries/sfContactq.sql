SELECT
  Contacts.contactID AS [Unique Id],
  Clients.clientID AS [ACCOUNT Unique Id],
  Contacts.contactLast AS LASTNAME,
  Contacts.contactFirst AS FIRSTNAME,
  Contacts.contactPrefix AS SALUTATION,
  Locations.LocationStreet AS OTHERSTREET,
  Locations.LCity AS OTHERCITY,
  Locations.LState AS OTHERSTATE,
  Locations.LZip AS OTHERPOSTALCODE,
  Locations.LCountry AS OTHERCOUNTRY,
  Null AS MAILINGSTREET,
  Null AS MAILINGCITY,
  Null AS MAILINGSTATE,
  Null AS MAILINGPOSTALCODE,
  Null AS MAILINGCOUNTRY,
  Contacts.contactPrimaryPhone AS PHONE,
  Contacts.contactFax AS FAX,
  Contacts.contactCell AS MOBILEPHONE,
  Null AS HOMEPHONE,
  Contacts.contactPhone2 AS OTHERPHONE,
  Contacts.ContactAsstPH AS ASSISTANTPHONE,
  0 AS REPORTSTO,
  Contacts.contactEmail AS EMAIL,
  Contacts.contactTitle AS TITLE,
  Contacts.contactClassification AS DEPARTMENT,
  [CONTACTASSTFNM] & "  " & [CONTACTASSTLNM] AS ASSISTANTNAME,
  Null AS LEADSOURCE,
  Null AS BIRTHDATE,
  Null AS DESCRIPTON,
  Null AS OWNER
FROM
  (
    Clients
    INNER JOIN Contacts ON Clients.clientID = Contacts.clientRef
  )
  LEFT JOIN Locations ON Contacts.cLocationID = Locations.LocationID
WHERE
  (
    (
      (Clients.clientType)= "client"
      Or (Clients.clientType)= "guarantor"
      Or (Clients.clientType) Like "*target"
    )
    AND (
      (Contacts.ContactObsolete)= No
    )
  )
ORDER BY
  Clients.clientID;
