SELECT
  Clients.ClientShNm,
  [schno] & [sgrpno] & [unitbuilddate] & [unitestdeldate] AS [Key],
  Schedule.SchNo,
  Schedule.SchFAID,
  SchGrp.SGrpNO,
  Units.UnitID,
  Units.unitvin,
  Units.UnitID,
  Right([unitvin], 6) AS VIn4,
  Units.unityr,
  Units.unitmake,
  Units.unittype,
  Units.unitstatus,
  Units.UnitEstDelDate,
  Units.UnitBuildDate,
  Units.UnitUnitNum,
  Schedule.SchShortDesc,
  Units.UnitAcceptDate,
  Units.UnitMSORecd,
  Units.UnitInvoiceNum,
  Units.UnitInvoiceDate,
  Units.UnitVendorPytDate,
  SchGrp.BLCD,
  Vendors.VendorName,
  Vendors.VendorContact,
  Vendors.VendorCity,
  Vendors.VendorState,
  Vendors.VendorZip,
  Vendors.VendorPhone,
  Vendors.VendorEmail,
  Units.UnitInServ,
  Units.UNITINSERVIND,
  Units.RevDel,
  Units.UnitDel,
  Units.delstatus
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
    RIGHT JOIN Units ON Vendors.VendorID = Units.VendorREF
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Clients.ClientShNm) Not Like "demo*"
    )
    AND (
      (Units.unitstatus)= "b"
    )
    AND (
      (MstrLease.MLOrig)= "FA"
    )
  )
ORDER BY
  Schedule.SchNo,
  SchGrp.SGrpNO,
  Units.unitvin,
  Right([unitvin], 6);
