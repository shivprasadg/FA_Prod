SELECT
  DISTINCT Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchNo,
  Schedule.SchID,
  Vendors.VendorName
FROM
  (
    ChildParts
    RIGHT JOIN (
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
      INNER JOIN Units ON SchGrp.SGrpSchID = Units.GroupREF
    ) ON ChildParts.UnitREF = Units.UnitID
  )
  LEFT JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID
WHERE
  (
    (
      (Vendors.VendorName) Is Not Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.AssignmentDate) Is Null
    )
  )
ORDER BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Schedule.SchID,
  Vendors.VendorName;
