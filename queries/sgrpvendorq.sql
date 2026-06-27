SELECT
  DISTINCT Clients.clientCompanyName,
  MstrLease.MLNo,
  Schedule.SchFAID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.SGrpFAID,
  SchGrp.SGrpID,
  Vendors.VendorID,
  Vendors.VendorName,
  Vendors.VENDORDBA,
  Vendors.VendorSNm,
  Vendors.VendorContact,
  Vendors.VendorAddress,
  Vendors.VendorCity,
  Vendors.VendorState,
  Vendors.VendorZip,
  Vendors.VendorPhone,
  Vendors.VendorFax,
  Vendors.VendorPytInstructions,
  Vendors.VendorEmail,
  Vendors.VendorPmtType,
  Vendors.VENDORPAYEE,
  Vendors.VENDORABA,
  Vendors.VENDORACH,
  Vendors.VENDORBKACCT,
  Vendors.VENDORBANK,
  Vendors.VENDORBADD1,
  Vendors.VEDNORBADD2,
  Vendors.VENDORBCITY,
  Vendors.VENDORBST,
  Vendors.VENDORBZIP,
  Vendors.VENDORPMTVER,
  Vendors.VendorObsolete,
  Vendors.VensorODt,
  Vendors.VendorTaxID,
  Units.UnitID,
  Units.UnitID
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    Vendors
    INNER JOIN Units ON Vendors.VendorID = Units.VendorREF
  ) ON SchGrp.SGrpID = Units.GroupREF
GROUP BY
  Clients.clientCompanyName,
  MstrLease.MLNo,
  Schedule.SchFAID,
  Schedule.SchNo,
  SchGrp.SGrpNO,
  SchGrp.SGrpFAID,
  SchGrp.SGrpID,
  Vendors.VendorID,
  Vendors.VendorName,
  Vendors.VENDORDBA,
  Vendors.VendorSNm,
  Vendors.VendorContact,
  Vendors.VendorAddress,
  Vendors.VendorCity,
  Vendors.VendorState,
  Vendors.VendorZip,
  Vendors.VendorPhone,
  Vendors.VendorFax,
  Vendors.VendorPytInstructions,
  Vendors.VendorEmail,
  Vendors.VendorPmtType,
  Vendors.VENDORPAYEE,
  Vendors.VENDORABA,
  Vendors.VENDORACH,
  Vendors.VENDORBKACCT,
  Vendors.VENDORBANK,
  Vendors.VENDORBADD1,
  Vendors.VEDNORBADD2,
  Vendors.VENDORBCITY,
  Vendors.VENDORBST,
  Vendors.VENDORBZIP,
  Vendors.VENDORPMTVER,
  Vendors.VendorObsolete,
  Vendors.VensorODt,
  Vendors.VendorTaxID,
  Units.UnitID,
  Units.UnitID,
  Clients.ClientShNm,
  MstrLease.MLOrig
HAVING
  (
    (
      (Clients.ClientShNm) Not Like "demo*"
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
  );
