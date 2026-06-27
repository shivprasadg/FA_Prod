SELECT
  Vendors.VendorName,
  Vendors.VendorID,
  INV_MsoContactq.contactID,
  INV_MsoContactq.VendorRef,
  INV_MsoContactq.contactEmail,
  INV_MsoContactq.VendorContact
FROM
  Vendors
  LEFT JOIN INV_MsoContactq ON Vendors.VendorID = INV_MsoContactq.VendorRef
WHERE
  (
    (
      (Vendors.VendorName)<> "Lessee"
      And (Vendors.VendorName)<> "unknown"
    )
  );
