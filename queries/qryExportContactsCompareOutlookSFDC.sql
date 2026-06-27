SELECT
  Contacts.contactPrefix AS Pre,
  Trim(
    [contactFirst] & IIf(
      Nz([contactMiddle], "")<> " ",
      " " & [contactMiddle],
      ""
    )
  ) AS FirstN,
  Trim(
    [contactLast] & IIf(
      Nz([contactSuffix], "")<> " ",
      " " & [contactSuffix],
      ""
    )
  ) AS LastN,
  [clientCompanyName] & IIf(
    Nz([ClientGroupName], "")<> "",
    " / " & [ClientGroupName],
    ""
  ) AS [Comp],
  Contacts.contactClassification AS Cls,
  Contacts.contactTitle AS Title,
  [LAddrLine1] & IIf(
    Nz([LAddrLine2], "")<> "",
    ", " & [LAddrLine2],
    ""
  )& IIf(
    Nz([LAddrLine3], "")<> "",
    ", " & [LAddrLine3],
    ""
  ) AS Address,
  Locations.LCity AS City,
  Locations.LState AS St,
  Locations.LZip AS Zip,
  Locations.LCountry AS Country,
  Null AS OthStr,
  Null AS OthCity,
  Null AS OthSt,
  Null AS OthZip,
  Contacts.ContactAsstPH AS AsstPhone,
  Contacts.contactFax AS Fax,
  [contactPrimaryPhone] & IIf(
    Nz([contactPhoneExt], "")<> " ",
    " ext. " & [contactPhoneExt],
    ""
  ) AS Phone,
  Null AS Home,
  Contacts.contactCell AS Mobile,
  Contacts.contactPhone2 AS OthPhone,
  Trim(
    IIf(
      Nz([ContactAsstFnm], "")<> "",
      [ContactAsstFnm],
      ""
    )& IIf(
      Nz([ContactAsstLnm], "")<> "",
      " " & [ContactAsstLnm],
      ""
    )
  ) AS AsstName,
  Null AS dob,
  Null AS cat,
  Contacts.contactemail AS Email,
  Null AS mgr,
  IIf(
    Nz([ContactObsolete], 0)= 1,
    "OBSOLETE ",
    ""
  )& IIf(
    Nz([DoNotCall], 0)= 1,
    "DO NOT CALL ",
    ""
  )& IIf(
    Nz([contactOptOut], 0)= 1,
    "OPT OUT ",
    ""
  )& [ContactNote] AS [Note],
  Contacts.contactID
FROM
  (
    (
      Contacts
      LEFT JOIN Clients ON Contacts.clientRef = Clients.clientID
    )
    LEFT JOIN Locations ON Contacts.cLocationID = Locations.LocationID
  )
  LEFT JOIN ClientGroups ON Clients.ClientGroupId = ClientGroups.ClientGroupId
ORDER BY
  Nz(
    [ClientGroupName], "zzzz" & [clientcompanyname]
  ),
  Contacts.contactLast,
  Contacts.contactFirst,
  Clients.ClientShNm,
  Contacts.contactEmail;
