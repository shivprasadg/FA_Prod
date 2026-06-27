SELECT
  Contacts.contactID,
  Contacts.LocationName,
  [Contacts]![contactAddress1] + " " + [Contacts]![contactCity] + ", " + [Contacts]![LState] AS address,
  Contacts.ClientREF,
  Contacts.contactClassification
FROM
  Contacts
WHERE
  (
    (
      (Contacts.ClientREF)= Forms!Units!clientID
    )
    And (
      (Contacts.contactClassification)= "Location"
    )
  )
ORDER BY
  Contacts.LocationName,
  [Contacts]![contactAddress1] + " " + [Contacts]![contactCity] + ", " + [Contacts]![LState];
