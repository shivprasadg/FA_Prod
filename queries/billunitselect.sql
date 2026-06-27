SELECT
  Units.*,
  Units.unitstatus,
  MstrLease.MLNo,
  Schedule.SchNo,
  Units.UnitVendorPytDate
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
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Units.unitstatus)<> "s"
      And (Units.unitstatus)<> "O"
      And (Units.unitstatus)<> "r"
    )
    AND (
      (Units.UnitDocAccpt)= Yes
    )
    AND (
      (Units.UnitInServ) Is Not Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.UnitAcceptDate) Is Not Null
    )
    AND (
      (Units.UnitInServInd)= "a"
    )
    AND (
      (Clients.ClientShNm) Not Like "demo*"
    )
  )
  OR (
    (
      (Units.unitstatus)<> "s"
      And (Units.unitstatus)<> "O"
      And (Units.unitstatus)<> "r"
    )
    AND (
      (Units.UnitDocAccpt)= No
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.UnitAcceptDate) Is Not Null
    )
    AND (
      (Clients.ClientShNm) Not Like "demo*"
    )
  )
  OR (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.UnitInServInd)= "a"
    )
    AND (
      (Clients.ClientShNm)= "toyota logistics"
    )
    AND (
      (SchGrp.BLCD) Is Not Null
    )
  )
  OR (
    (
      (Units.unitstatus)<> "s"
      And (Units.unitstatus)<> "O"
      And (Units.unitstatus)<> "r"
    )
    AND (
      (Units.UnitDocAccpt)= Yes
    )
    AND (
      (Units.UnitInServ) Is Null
    )
    AND (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (Units.UnitAcceptDate) Is Not Null
    )
    AND (
      (Clients.ClientShNm) Not Like "demo*"
    )
    AND (
      (Units.UnitPDStartDate) Is Not Null
    )
  );
