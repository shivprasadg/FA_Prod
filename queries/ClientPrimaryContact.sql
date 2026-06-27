SELECT
  Contacts.*,
  Contacts.contactID
FROM
  Contacts
WHERE
  (
    (
      (Contacts.contactClassification)= "Primary"
    )
  );
