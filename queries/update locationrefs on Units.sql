SELECT
  Units.UnitID,
  Units.UnitRegLocationRef,
  Contacts.oldLocID,
  Contacts.contactID
FROM
  Units
  INNER JOIN Contacts ON Units.UnitRegLocationRef = Contacts.oldLocID;
