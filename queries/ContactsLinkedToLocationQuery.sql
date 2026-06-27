SELECT
  ContactsLinkedToLocation.ContactID,
  ContactsLinkedToLocation.LocationID,
  ContactsLinkedToLocationTypes.LinkType,
  [contactFirst] & " " & [contactLast] AS Contact,
  Contacts.contactEmail,
  Locations.Locationsn,
  Locations.LCity AS City,
  Locations.LState AS State,
  Locations.LocationType,
  ContactsLinkedToLocation.LinkTypeID
FROM
  (
    (
      (
        ContactsLinkedToLocation
        INNER JOIN Contacts ON ContactsLinkedToLocation.ContactID = Contacts.contactID
      )
      INNER JOIN Locations ON ContactsLinkedToLocation.LocationID = Locations.LocationID
    )
    INNER JOIN Clients ON Locations.ClientREF = Clients.clientID
  )
  INNER JOIN ContactsLinkedToLocationTypes ON ContactsLinkedToLocation.LinkTypeID = ContactsLinkedToLocationTypes.LinkTypeID;
