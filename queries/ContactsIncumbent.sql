SELECT
  Contacts.clientRef,
  Contacts.contactTitle AS ClientSignTitle,
  Contacts.contactFullName AS ClientSignFullName,
  Contacts.contactID
FROM
  Contacts
WHERE
  (
    (
      (Contacts.contactInc)<> No
    )
    AND (
      (Contacts.ContactObsolete)<> Yes
    )
  );
