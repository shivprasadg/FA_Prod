SELECT
  VendorContacts.*
FROM
  Contacts,
  VendorContacts
WHERE
  (
    (
      (VendorContacts.VendorRef)= [forms]![vendor form]![vendorid]
    )
  );
