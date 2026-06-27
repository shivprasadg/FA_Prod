SELECT
  Vendors.VendorSNm,
  Vendors.VendorName,
  Vendors.VendorAddress,
  Vendors.VendorCity,
  Vendors.VendorState,
  Vendors.VendorZip,
  VendorContacts.contactLast AS LastName,
  VendorContacts.contactFirst AS FristName,
  VendorContacts.contactMiddle AS Middle,
  VendorContacts.contactTitle AS Title,
  VendorContacts.contactPrimaryPhone AS Phone,
  VendorContacts.contactEmail AS Email,
  VendorContacts.contactCell AS Mobil,
  VendorContacts.contactAddress1 AS Address1,
  VendorContacts.contactAddress2 AS Addrees2,
  VendorContacts.contactCity AS City,
  VendorContacts.LState AS State,
  VendorContacts.contactZip AS Zip,
  Vendors.VENDORBANK,
  IIf([VCEMoptout] = False, "IN", "OUT") AS OPT_In_Out
FROM
  Vendors
  LEFT JOIN VendorContacts ON Vendors.VendorID = VendorContacts.VendorRef
WHERE
  (
    (
      (VendorContacts.ContactObsolete)= No
    )
  )
ORDER BY
  Vendors.VendorSNm,
  VendorContacts.contactLast,
  VendorContacts.contactFirst;
