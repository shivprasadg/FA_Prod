SELECT
  Itinerary.FrDt,
  Itinerary.FrTime,
  Itinerary.ToTime,
  Itinerary.TimeZone,
  Itinerary.Activity,
  Itinerary.Comment,
  IIf(
    IsNull([city]),
    IIf(
      IsNull([LCity]),
      [clientcity],
      [LCity]
    ),
    [city]
  )& " " & IIf(
    IsNull([state]),
    IIf(
      IsNull([locations].[LState]),
      [clientstate],
      [locations].[LState]
    ),
    [state]
  ) AS CityST,
  Salesperson.SPInitials,
  Itinerary.LOEXRef,
  "Others: " & IIf([ab] = Yes, "AB ", Null)& IIf([mm] = Yes, "MM ", Null)& IIf(
    [sm] = Yes,
    "SM ",
    IIf([pf] = Yes, "PF ", Null)& IIf([JF] = Yes, " JF ", Null)& IIf([du] = Yes, "DU ", Null)
  ) AS Others,
  Itinerary.Allday,
  Itinerary.ItinID,
  Itinerary.City,
  Itinerary.State,
  Itinerary.Allday,
  Itinerary.Contact,
  Itinerary.ClientProsp,
  Clients.ClientShNm,
  Clients.clientCity,
  Clients.clientState,
  Contacts.contactFullName,
  Locations.LCity,
  Locations.LState,
  Clients.creditFAProfile,
  Clients.creditFAScore,
  Clients.creditRank,
  Contacts.contacttitle
FROM
  (
    (
      (
        Salesperson
        INNER JOIN Itinerary ON Salesperson.SalespersonID = Itinerary.LOEXRef
      )
      LEFT JOIN Contacts ON Itinerary.Contact = Contacts.contactID
    )
    LEFT JOIN Locations ON Contacts.cLocationID = Locations.LocationID
  )
  LEFT JOIN Clients ON Itinerary.ClientProsp = Clients.clientID
WHERE
  (
    (
      (Itinerary.FrDt)>= [forms]![itinupdtF]![itinstart]
      And (Itinerary.FrDt)<= [forms]![itinupdtf]![itinend]
    )
  );
