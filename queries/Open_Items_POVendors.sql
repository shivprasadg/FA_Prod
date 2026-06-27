SELECT
  DISTINCT Tracking_Open_Units_Clients.FAID,
  Tracking_Open_Units_Clients.VendorID,
  Tracking_Open_Units_Clients.Client,
  Tracking_Open_Units_Clients.clientCompanyName AS ClientName,
  SchGrp.SGrpNO,
  Nz([PONum1], "TBD") AS PO,
  SchGrp.POVendor,
  Vendors.VendorName,
  Tracking_Open_Units_Clients.Units
FROM
  (
    Tracking_Open_Units_Clients
    INNER JOIN SchGrp ON Tracking_Open_Units_Clients.FAID = SchGrp.SGrpFAID
  )
  INNER JOIN Vendors ON Tracking_Open_Units_Clients.VendorID = Vendors.VendorID
GROUP BY
  Tracking_Open_Units_Clients.FAID,
  Tracking_Open_Units_Clients.VendorID,
  Tracking_Open_Units_Clients.Client,
  Tracking_Open_Units_Clients.clientCompanyName,
  SchGrp.SGrpNO,
  Nz([PONum1], "TBD"),
  SchGrp.POVendor,
  Vendors.VendorName,
  Tracking_Open_Units_Clients.Units
HAVING
  (
    (
      (SchGrp.POVendor) Is Not Null
    )
  );
