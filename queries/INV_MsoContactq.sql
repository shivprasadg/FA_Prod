SELECT
  VendorContacts.contactID,
  VendorContacts.VendorRef,
  VendorContacts.contactEmail,
  [contactfirst] & [contactmiddle] & " " & [contactlast] AS VendorContact
FROM
  VendorContacts
WHERE
  (
    (
      (VendorContacts.INV_MSO)= True
    )
  );
