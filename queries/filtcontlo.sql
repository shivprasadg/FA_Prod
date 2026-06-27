SELECT
  Contacts.*
FROM
  Contacts
WHERE
  (
    (
      (Contacts.contactID)= [forms]![searchclients_sublo]![contactid]
    )
  );
