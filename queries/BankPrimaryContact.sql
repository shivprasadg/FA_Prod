SELECT
  Contacts.*,
  Contacts.contactID AS BankID
FROM
  Contacts
WHERE
  (
    (
      (Contacts.contactClassification)= "Primary"
    )
  );
